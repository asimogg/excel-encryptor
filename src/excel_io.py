"""Excel okuma / şifreleyip yazma / çözüp yazma katmanı."""

from __future__ import annotations

import base64
import json
import os

# Excel dosyaları dışarıdan gelebilir. openpyxl, kötü niyetli XML'e (XXE,
# "billion laughs" varlık genişletmesi) karşı kendi başına koruma sağlamaz;
# defusedxml standart XML çözümleyicisini güvenli hale getirir.
try:
    import defusedxml
    defusedxml.defuse_stdlib()
except ImportError:                                  # pragma: no cover
    defusedxml = None

import openpyxl
from openpyxl.utils import get_column_letter

import crypto_engine as ce

# Kaynak tüketimi saldırılarına (zip bomb benzeri) karşı üst sınırlar
MAX_CELLS = 5_000_000


class ExcelError(Exception):
    pass


# --------------------------------------------------------------------------- #
# Yükleme
# --------------------------------------------------------------------------- #

def load_workbook(path: str):
    if not os.path.isfile(path):
        raise ExcelError(f"Dosya bulunamadı: {path}")
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise ExcelError(f"Excel dosyası açılamadı: {exc}") from exc

    cells = sum((ws.max_row or 0) * (ws.max_column or 0) for ws in wb.worksheets)
    if cells > MAX_CELLS:
        raise ExcelError(
            f"Dosya çok büyük ({cells:,} hücre). Bellek tükenmesini önlemek için "
            f"{MAX_CELLS:,} hücre sınırı var.".replace(",", "."))
    return wb


def read_meta(wb) -> dict | None:
    """Dosya daha önce bu programla şifrelendiyse metadata'sını döndürür."""
    if ce.META_SHEET not in wb.sheetnames:
        return None
    raw = wb[ce.META_SHEET].cell(row=1, column=1).value
    if not raw:
        return None
    try:
        return json.loads(raw)
    except Exception:  # noqa: BLE001
        return None


def data_sheet_names(wb) -> list[str]:
    return [n for n in wb.sheetnames if n != ce.META_SHEET]


def grid(ws, max_rows: int = 5000, max_cols: int = 200):
    """Sayfayı ekranda göstermek için 2B listeye çevirir.

    Dönen: (satırlar, gösterilen_satır_sayısı, gösterilen_kolon_sayısı,
            toplam_satır, toplam_kolon)
    """
    total_rows = ws.max_row or 1
    total_cols = ws.max_column or 1
    n_rows = min(total_rows, max_rows)
    n_cols = min(total_cols, max_cols)

    rows = []
    for r in range(1, n_rows + 1):
        row = []
        for c in range(1, n_cols + 1):
            v = ws.cell(row=r, column=c).value
            row.append("" if v is None else str(v))
        rows.append(row)
    return rows, n_rows, n_cols, total_rows, total_cols


def column_headers(n_cols: int, first_row=None) -> list[str]:
    """'A · Ad Soyad' biçiminde başlıklar — hem Excel harfi hem ilk satır metni."""
    out = []
    for c in range(n_cols):
        letter = get_column_letter(c + 1)
        label = ""
        if first_row and c < len(first_row):
            label = str(first_row[c] or "").strip()
        out.append(f"{letter} · {label}" if label else letter)
    return out


def scan_encrypted_cells(wb) -> dict[str, list[tuple[int, int]]]:
    """Dosyadaki tüm şifreli hücreleri bulur (metadata'ya güvenmeden)."""
    found: dict[str, list[tuple[int, int]]] = {}
    for name in data_sheet_names(wb):
        ws = wb[name]
        hits = [
            (cell.row, cell.column)
            for row in ws.iter_rows()
            for cell in row
            if ce.is_token(cell.value)
        ]
        if hits:
            found[name] = hits
    return found


# --------------------------------------------------------------------------- #
# Çıktı yolu
# --------------------------------------------------------------------------- #

def output_path(src: str, suffix: str) -> str:
    """'.../Rapor.xlsx' + 'encrypted' -> '.../Rapor_encrypted.xlsx'

    Zaten '_encrypted' ile bitiyorsa ek üst üste binmez.
    """
    folder, filename = os.path.split(src)
    stem, ext = os.path.splitext(filename)
    for known in ("_encrypted", "_decrypted"):
        if stem.endswith(known):
            stem = stem[: -len(known)]
    if ext.lower() not in (".xlsx", ".xlsm"):
        ext = ".xlsx"
    return os.path.join(folder, f"{stem}_{suffix}{ext}")


def unique_path(path: str) -> str:
    """Aynı isimde dosya varsa '(1)', '(2)' ... ekler."""
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{stem} ({i}){ext}"):
        i += 1
    return f"{stem} ({i}){ext}"


# --------------------------------------------------------------------------- #
# Şifrele / Çöz
# --------------------------------------------------------------------------- #

def _number_format_for(value) -> str:
    """Çözülen değerin Excel'de doğru tipte görünmesi için biçim kodu.

    Tarih/saat hücrelerinde biçim 'General' bırakılırsa Excel (ve openpyxl)
    değeri seri numara olarak okur; bu yüzden tipe göre biçim veriyoruz.
    """
    import datetime as dt

    if isinstance(value, dt.datetime):
        return "yyyy-mm-dd hh:mm:ss"
    if isinstance(value, dt.date):
        return "yyyy-mm-dd"
    if isinstance(value, dt.time):
        return "hh:mm:ss"
    return "General"


def encrypt_workbook(src_path, selections, method_id, password, dest_path=None, progress=None):
    """Seçili hücreleri şifreleyip yeni bir dosyaya yazar.

    selections: {sayfa_adı: {(satır, kolon), ...}}  (Excel 1-tabanlı)
    Dönen: (yazılan_yol, şifrelenen_hücre_sayısı)
    """
    method = ce.METHODS_BY_ID.get(method_id)
    if method is None:
        raise ExcelError(f"Bilinmeyen yöntem: {method_id}")
    problem = ce.password_problem(password or "")
    if problem:
        raise ExcelError(problem)

    total = sum(len(v) for v in selections.values())
    if total == 0:
        raise ExcelError("Şifrelenecek hücre seçilmedi.")

    wb = load_workbook(src_path)
    if read_meta(wb):
        raise ExcelError("Bu dosya zaten şifrelenmiş görünüyor. Önce çözün.")

    salt = ce.new_salt()
    key = ce.derive_key(password, salt)

    done = 0
    written_count = 0
    for sheet_name, cells in selections.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for (r, c) in sorted(cells):
            cell = ws.cell(row=r, column=c)
            if cell.value is None or ce.is_token(cell.value):
                done += 1
                continue
            cell.value = ce.encrypt_value(cell.value, method_id, key)
            cell.number_format = "@"  # metin — Excel sayıya çevirmeye çalışmasın
            written_count += 1
            done += 1
            if progress and done % 250 == 0:
                progress(done, total)

    if not written_count:
        raise ExcelError("Seçilen hücrelerin hepsi boştu — şifrelenecek veri yok.")

    # Yalnızca çözmek için gereken asgari bilgi tutulur: anahtar doğrulama
    # etiketi (parola tahmin oracle'ı) ve hangi hücrelerin gizlendiğini
    # gösteren harita bilinçli olarak yazılmaz.
    meta = {
        "version": ce.META_VERSION,
        "method": method_id,
        "kdf": "pbkdf2-sha256",
        "iterations": ce.KDF_ITERATIONS,
        "salt": base64.b64encode(salt).decode(),
    }
    ms = wb.create_sheet(ce.META_SHEET)
    ms.cell(row=1, column=1).value = json.dumps(meta, ensure_ascii=False)
    ms.cell(row=3, column=1).value = (
        "Bu sayfa Excel Encryptor tarafından oluşturuldu. Silmeyin — "
        "şifre çözme bilgileri burada tutulur. Anahtarın kendisi burada YOKTUR."
    )
    ms.sheet_state = "hidden"

    dest = unique_path(dest_path or output_path(src_path, "encrypted"))
    try:
        wb.save(dest)
    except Exception as exc:  # noqa: BLE001
        raise ExcelError(f"Dosya kaydedilemedi: {exc}") from exc

    if progress:
        progress(total, total)
    return dest, written_count


def decrypt_workbook(src_path, password, dest_path=None, progress=None):
    """Şifreli dosyayı çözüp yeni bir dosyaya yazar.

    Dönen: (yazılan_yol, çözülen_hücre_sayısı)
    """
    wb = load_workbook(src_path)
    meta = read_meta(wb)
    found = scan_encrypted_cells(wb)

    if not found:
        raise ExcelError("Bu dosyada şifrelenmiş hücre bulunamadı.")

    if not meta:
        raise ExcelError(
            "Dosyanın şifre bilgileri (gizli _ENCMETA_ sayfası) silinmiş; "
            "bu dosya çözülemez."
        )
    if not password:
        raise ExcelError("Anahtar boş olamaz.")

    try:
        salt = base64.b64decode(meta["salt"], validate=True)
        iterations = int(meta.get("iterations", ce.KDF_ITERATIONS))
    except (KeyError, ValueError, TypeError) as exc:
        raise ExcelError("Dosyanın şifre bilgileri bozuk; bu dosya çözülemez.") from exc
    key = ce.derive_key(password, salt, iterations)

    # Anahtarı, ilk şifreli hücreyi çözmeyi deneyerek doğruluyoruz. Dosyada
    # ayrı bir doğrulama etiketi tutmuyoruz; o etiket saldırgana anahtarı
    # veriye dokunmadan deneme imkânı verirdi.
    first_sheet = next(iter(found))
    first_cell = found[first_sheet][0]
    try:
        ce.decrypt_value(wb[first_sheet].cell(*first_cell).value, key)
    except ce.CryptoError as exc:
        raise ExcelError(
            "Anahtar hatalı. Şifreleme sırasında girilen anahtarı girin."
        ) from exc

    total = sum(len(v) for v in found.values())
    done = 0
    for sheet_name, cells in found.items():
        ws = wb[sheet_name]
        for (r, c) in cells:
            cell = ws.cell(row=r, column=c)
            value = ce.decrypt_value(cell.value, key)
            cell.value = value
            cell.number_format = _number_format_for(value)
            done += 1
            if progress and done % 250 == 0:
                progress(done, total)

    del wb[ce.META_SHEET]

    dest = unique_path(dest_path or output_path(src_path, "decrypted"))
    try:
        wb.save(dest)
    except Exception as exc:  # noqa: BLE001
        raise ExcelError(f"Dosya kaydedilemedi: {exc}") from exc

    if progress:
        progress(total, total)
    return dest, total
