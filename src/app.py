#!/usr/bin/env python3
"""Excel Encryptor — Excel dosyalarındaki gizli verileri hücre bazında şifreler.

Çalıştırma:  python src/app.py
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tksheet import Sheet

import crypto_engine as ce
import excel_io as xio

APP_TITLE = "Excel Encryptor"
MARK_BG = "#ffb85c"
MARK_FG = "#1a1a1a"
ENC_BG = "#8ec9ff"
ENC_FG = "#0b2b46"

# Windows'taki Tk 8.6 renkli emoji glyph'lerini güvenilir çizemiyor; oradaki
# etiketlerden emojileri düşürüyoruz (geometrik semboller her yerde çalışır).
_WIN = os.name == "nt"


def L(text: str) -> str:
    """Etiketi çalışılan platformda düzgün görünecek hale getirir."""
    if not _WIN:
        return text
    for emoji in ("📂", "🔒", "🔓", "🔎"):
        text = text.replace(emoji + "  ", "").replace(emoji + " ", "").replace(emoji, "")
    return text.replace("⌫", "×")


# Platforma uygun arayüz yazı tipi ("Helvetica" Windows'ta gerçek bir font değil)
UI_FONT = "Segoe UI" if os.name == "nt" else "Helvetica"

MAX_DISPLAY_ROWS = 5000
MAX_DISPLAY_COLS = 200

# "Hassas kolonları öner" için başlık anahtar kelimeleri
SENSITIVE_PATTERNS = [
    r"\btc\b", r"tckn", r"kimlik", r"\bt\.?c\.? ?no", r"vergi ?no", r"vkn",
    r"ad[ıi]? ?soyad", r"\bisim\b", r"\bad\b", r"soyad", r"\bname\b", r"surname",
    r"telefon", r"\bgsm\b", r"cep", r"phone", r"\bmobil",
    r"e-?mail", r"e-?posta", r"\bmail\b",
    r"adres", r"address",
    r"iban", r"hesap ?no", r"kart ?no", r"card",
    r"maa[şs]", r"[üu]cret", r"salary", r"wage", r"prim",
    r"do[ğg]um", r"birth",
    r"parola", r"[şs]ifre", r"password", r"secret", r"token", r"api ?key",
    r"sicil", r"\bssn\b", r"pasaport", r"passport",
    r"form[uü]l", r"re[çc]ete", r"formulation",
]
SENSITIVE_RE = re.compile("|".join(SENSITIVE_PATTERNS), re.IGNORECASE)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x760")
        self.minsize(1040, 620)

        # --- durum ---
        self.src_path: str | None = None
        self.wb = None
        self.meta = None                       # şifreli dosyaysa metadata
        self.sheet_names: list[str] = []
        self.current_sheet: str | None = None
        self.marks: dict[str, set[tuple[int, int]]] = {}   # sayfa -> {(satır, kolon)} 1-tabanlı
        self.enc_cells: dict[str, set[tuple[int, int]]] = {}
        self.sheet_dims: dict[str, tuple[int, int]] = {}   # sayfa -> (toplam_satır, toplam_kolon)
        self.busy = False

        self._build_ui()
        self._refresh_method_note()
        self._set_status("Başlamak için bir Excel dosyası yükleyin.")

    # ------------------------------------------------------------------ #
    # Arayüz
    # ------------------------------------------------------------------ #
    def _build_ui(self):
        style = ttk.Style(self)
        for theme in (("vista", "winnative") if _WIN else ("clam",)):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        style.configure("Accent.TButton", font=(UI_FONT, 12, "bold"), padding=8)
        style.configure("Big.TButton", font=(UI_FONT, 11), padding=6)
        style.configure("Hint.TLabel", foreground="#666")
        style.configure("Head.TLabel", font=(UI_FONT, 11, "bold"))

        # --- üst şerit ---
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")

        ttk.Button(top, text=L("📂  Excel Yükle"), style="Big.TButton",
                   command=self.on_load).pack(side="left")

        self.file_var = tk.StringVar(value="— dosya seçilmedi —")
        ttk.Label(top, textvariable=self.file_var).pack(side="left", padx=12)

        self.sheet_combo = ttk.Combobox(top, state="disabled", width=26, values=[])
        self.sheet_combo.pack(side="right")
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)
        ttk.Label(top, text="Sayfa:").pack(side="right", padx=(0, 6))

        ttk.Separator(self).pack(fill="x")

        # --- gövde ---
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        self.sheet = Sheet(
            body,
            data=[[""]],
            headers=["A"],
            show_row_index=True,
            header_font=(UI_FONT, 11, "normal"),
            font=(UI_FONT, 11, "normal"),
            empty_horizontal=0,
            empty_vertical=0,
        )
        self.sheet.enable_bindings(
            "single_select", "drag_select", "column_select", "row_select",
            "ctrl_select", "arrowkeys", "column_width_resize",
            "double_click_column_resize", "copy", "select_all",
        )
        self.sheet.pack(side="left", fill="both", expand=True, padx=(10, 6), pady=8)
        self.sheet.extra_bindings("cell_select", lambda e: None)
        self.sheet.bind("<Double-Button-1>", self._on_double_click, add="+")

        side = ttk.Frame(body, padding=(6, 8, 12, 8))
        side.pack(side="right", fill="y")
        self._build_side(side)

        # --- durum çubuğu ---
        bar = ttk.Frame(self, padding=(10, 6))
        bar.pack(fill="x")
        self.status_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.status_var).pack(side="left")
        self.progress = ttk.Progressbar(bar, length=180, mode="determinate")
        self.progress.pack(side="right")

    def _build_side(self, side):
        """Sağ panel: iki sekme — şifreleme ve şifre çözme."""
        self.nb = ttk.Notebook(side)
        self.nb.pack(fill="both", expand=True)

        enc_tab = ttk.Frame(self.nb, padding=(8, 10))
        dec_tab = ttk.Frame(self.nb, padding=(8, 10))
        self.nb.add(enc_tab, text=L("  🔒  Şifrele  "))
        self.nb.add(dec_tab, text=L("  🔓  Şifre Çöz  "))

        self._build_encrypt_tab(enc_tab)
        self._build_decrypt_tab(dec_tab)

    def _build_encrypt_tab(self, side):
        w = 30

        ttk.Label(side, text="1 · Gizlenecek Veriyi Seç", style="Head.TLabel").pack(anchor="w")
        ttk.Label(side, style="Hint.TLabel", wraplength=250, justify="left",
                  text="Tabloda hücre / satır / kolon seçin, sonra işaretleyin. "
                       "Tek hücreye çift tıklamak da işareti açıp kapatır.").pack(anchor="w", pady=(2, 8))

        ttk.Button(side, text=L("✓  Seçileni İşaretle"), width=w,
                   command=self.mark_selection).pack(fill="x", pady=2)
        ttk.Button(side, text=L("✕  Seçimin İşaretini Kaldır"), width=w,
                   command=lambda: self.mark_selection(unmark=True)).pack(fill="x", pady=2)
        ttk.Button(side, text=L("▐  Kolonun Tamamını İşaretle"), width=w,
                   command=self.mark_full_columns).pack(fill="x", pady=2)
        ttk.Button(side, text=L("▬  Satırın Tamamını İşaretle"), width=w,
                   command=self.mark_full_rows).pack(fill="x", pady=2)
        ttk.Button(side, text=L("🔎  Hassas Kolonları Öner"), width=w,
                   command=self.suggest_sensitive).pack(fill="x", pady=2)
        ttk.Button(side, text=L("⌫  Bu Sayfadaki İşaretleri Sil"), width=w,
                   command=self.clear_sheet_marks).pack(fill="x", pady=2)
        ttk.Button(side, text=L("⌫⌫  Tüm İşaretleri Sil"), width=w,
                   command=self.clear_all_marks).pack(fill="x", pady=2)

        self.count_var = tk.StringVar(value="işaretli hücre: 0")
        ttk.Label(side, textvariable=self.count_var, style="Head.TLabel").pack(anchor="w", pady=(10, 0))

        ttk.Separator(side).pack(fill="x", pady=12)

        ttk.Label(side, text="2 · Yöntem ve Anahtar", style="Head.TLabel").pack(anchor="w")
        self.method_combo = ttk.Combobox(side, state="readonly", width=w - 2,
                                         values=[m.label for m in ce.METHODS])
        self.method_combo.current(0)
        self.method_combo.pack(fill="x", pady=(6, 2))
        self.method_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_method_note())

        self.method_note = ttk.Label(side, style="Hint.TLabel", wraplength=250, justify="left")
        self.method_note.pack(anchor="w", pady=(0, 8))

        self.key_var = tk.StringVar()
        self.key2_var = tk.StringVar()
        self.show_key = tk.BooleanVar(value=False)

        ttk.Label(side, text="Anahtar (parola)").pack(anchor="w")
        self.key_entry = ttk.Entry(side, textvariable=self.key_var, show="•", width=w - 2)
        self.key_entry.pack(fill="x", pady=(2, 4))

        ttk.Label(side, text="Anahtar (tekrar)").pack(anchor="w")
        self.key2_entry = ttk.Entry(side, textvariable=self.key2_var, show="•", width=w - 2)
        self.key2_entry.pack(fill="x", pady=(2, 4))

        ttk.Checkbutton(side, text="Anahtarı göster", variable=self.show_key,
                        command=self._toggle_key_visibility).pack(anchor="w")

        ttk.Separator(side).pack(fill="x", pady=12)

        ttk.Label(side, text="3 · Uygula", style="Head.TLabel").pack(anchor="w")
        self.enc_btn = ttk.Button(side, text=L("🔒  ŞİFRELE"), style="Accent.TButton",
                                  command=self.on_encrypt)
        self.enc_btn.pack(fill="x", pady=(6, 4))

        ttk.Label(side, style="Hint.TLabel", wraplength=250, justify="left",
                  text="Çıktı, kaynak dosyayla aynı klasöre "
                       "Dosya_adı_encrypted.xlsx olarak yazılır. "
                       "Orijinal dosyaya dokunulmaz.").pack(anchor="w", pady=(8, 0))

    def _build_decrypt_tab(self, side):
        w = 30

        ttk.Label(side, text="1 · Şifreli Dosyayı Aç", style="Head.TLabel").pack(anchor="w")
        ttk.Label(side, style="Hint.TLabel", wraplength=250, justify="left",
                  text="Üstteki “Excel Yükle” ile daha önce bu programla şifrelenmiş "
                       "bir dosya seçin. Şifreli hücreler tabloda mavi görünür.").pack(anchor="w", pady=(2, 8))

        ttk.Button(side, text=L("📂  Şifreli Excel Yükle"), width=w,
                   command=self.on_load).pack(fill="x", pady=2)

        self.dec_status_var = tk.StringVar(value="Henüz dosya yüklenmedi.")
        ttk.Label(side, textvariable=self.dec_status_var, style="Head.TLabel",
                  wraplength=250, justify="left").pack(anchor="w", pady=(12, 0))

        self.dec_detail_var = tk.StringVar(value="")
        ttk.Label(side, textvariable=self.dec_detail_var, style="Hint.TLabel",
                  wraplength=250, justify="left").pack(anchor="w", pady=(2, 0))

        ttk.Separator(side).pack(fill="x", pady=12)

        ttk.Label(side, text="2 · Anahtar", style="Head.TLabel").pack(anchor="w")
        ttk.Label(side, style="Hint.TLabel", wraplength=250, justify="left",
                  text="Şifreleme sırasında girilen anahtarın aynısı. "
                       "Yöntem dosyadan otomatik okunur.").pack(anchor="w", pady=(2, 6))

        self.dkey_var = tk.StringVar()
        self.show_dkey = tk.BooleanVar(value=False)
        self.dkey_entry = ttk.Entry(side, textvariable=self.dkey_var, show="•", width=w - 2)
        self.dkey_entry.pack(fill="x", pady=(2, 4))
        self.dkey_entry.bind("<Return>", lambda e: self.on_decrypt())

        ttk.Checkbutton(side, text="Anahtarı göster", variable=self.show_dkey,
                        command=self._toggle_key_visibility).pack(anchor="w")

        ttk.Separator(side).pack(fill="x", pady=12)

        ttk.Label(side, text="3 · Uygula", style="Head.TLabel").pack(anchor="w")
        self.dec_btn = ttk.Button(side, text=L("🔓  ŞİFREYİ ÇÖZ"), style="Accent.TButton",
                                  command=self.on_decrypt)
        self.dec_btn.pack(fill="x", pady=(6, 4))

        ttk.Label(side, style="Hint.TLabel", wraplength=250, justify="left",
                  text="Çözülmüş kopya, şifreli dosyayla aynı klasöre "
                       "Dosya_adı_decrypted.xlsx olarak yazılır. "
                       "Şifreli dosyaya dokunulmaz.").pack(anchor="w", pady=(8, 0))

    def _refresh_decrypt_panel(self):
        """Yüklü dosyanın şifre durumunu çözme sekmesine yansıtır."""
        n_enc = sum(len(v) for v in self.enc_cells.values())
        if not self.src_path:
            self.dec_status_var.set("Henüz dosya yüklenmedi.")
            self.dec_detail_var.set("")
        elif not n_enc:
            self.dec_status_var.set("Bu dosyada şifreli hücre yok.")
            self.dec_detail_var.set(os.path.basename(self.src_path))
        else:
            mid = (self.meta or {}).get("method", "")
            label = ce.METHODS_BY_ID[mid].label if mid in ce.METHODS_BY_ID else "bilinmiyor"
            sheets = ", ".join(self.enc_cells)
            self.dec_status_var.set(
                L(f"🔒 {n_enc:,} şifreli hücre bulundu.").replace(",", "."))
            self.dec_detail_var.set(f"Dosya: {os.path.basename(self.src_path)}\n"
                                    f"Yöntem: {label}\nSayfalar: {sheets}")

    def _toggle_key_visibility(self):
        show = "" if self.show_key.get() else "•"
        self.key_entry.configure(show=show)
        self.key2_entry.configure(show=show)
        self.dkey_entry.configure(show="" if self.show_dkey.get() else "•")

    def _refresh_method_note(self):
        m = ce.METHODS_BY_LABEL[self.method_combo.get()]
        text = m.note
        if m.warning:
            text += "\n\n⚠ " + m.warning
        self.method_note.configure(text=text)

    def _set_status(self, text):
        self.status_var.set(text)
        self.update_idletasks()

    # ------------------------------------------------------------------ #
    # Dosya yükleme
    # ------------------------------------------------------------------ #
    def on_load(self):
        path = filedialog.askopenfilename(
            title="Excel dosyası seç",
            filetypes=[("Excel dosyaları", "*.xlsx *.xlsm"), ("Tüm dosyalar", "*.*")],
        )
        if not path:
            return
        self._set_status("Dosya okunuyor…")
        try:
            wb = xio.load_workbook(path)
        except xio.ExcelError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            self._set_status("Dosya açılamadı.")
            return

        self.src_path = path
        self.wb = wb
        self.meta = xio.read_meta(wb)
        self.sheet_names = xio.data_sheet_names(wb)
        self.marks = {n: set() for n in self.sheet_names}
        self.sheet_dims = {}

        found = xio.scan_encrypted_cells(wb)
        self.enc_cells = {k: set(v) for k, v in found.items()}

        self.file_var.set(os.path.basename(path))
        self.sheet_combo.configure(state="readonly", values=self.sheet_names)
        if self.sheet_names:
            self.sheet_combo.current(0)
            self.show_sheet(self.sheet_names[0])

        self._refresh_decrypt_panel()

        n_enc = sum(len(v) for v in self.enc_cells.values())
        if n_enc:
            method = (self.meta or {}).get("method", "?")
            label = ce.METHODS_BY_ID[method].label if method in ce.METHODS_BY_ID else method
            self._set_status(f"Şifreli dosya · {n_enc} hücre · yöntem: {label} — anahtarı girip ÇÖZ'e basın.")
            if self.meta and method in ce.METHODS_BY_ID:
                self.method_combo.set(ce.METHODS_BY_ID[method].label)
                self._refresh_method_note()
            self.nb.select(1)            # şifreli dosya → çözme sekmesi
            self.dkey_entry.focus_set()
        else:
            self._set_status(f"{len(self.sheet_names)} sayfa yüklendi. Gizlenecek hücreleri seçin.")
            self.nb.select(0)

    def on_sheet_change(self, _event=None):
        self.show_sheet(self.sheet_combo.get())

    def show_sheet(self, name):
        if not self.wb or name not in self.wb.sheetnames:
            return
        self.current_sheet = name
        ws = self.wb[name]
        rows, n_rows, n_cols, total_rows, total_cols = xio.grid(
            ws, MAX_DISPLAY_ROWS, MAX_DISPLAY_COLS
        )
        self.sheet_dims[name] = (total_rows, total_cols)

        display = [[self._preview(v) for v in row] for row in rows]
        headers = xio.column_headers(n_cols, rows[0] if rows else None)

        self.sheet.set_sheet_data(display or [[""]], reset_highlights=True)
        self.sheet.headers(headers or ["A"])
        self.sheet.set_all_column_widths()
        self._redraw_marks()

        note = ""
        if total_rows > n_rows or total_cols > n_cols:
            note = (f"  (ekranda ilk {n_rows}×{n_cols} gösteriliyor; "
                    f"dosyada {total_rows}×{total_cols} — kolon/satır işaretleme tamamına uygulanır)")
        self._set_status(f"Sayfa: {name} · {total_rows} satır × {total_cols} kolon{note}")

    @staticmethod
    def _preview(value: str) -> str:
        """Şifreli token'lar tabloda kısaltılmış gösterilir."""
        if ce.is_token(value):
            return ("# " if _WIN else "🔒 ") + value.split(":", 2)[2][:14] + "…"
        return value

    # ------------------------------------------------------------------ #
    # İşaretleme
    # ------------------------------------------------------------------ #
    def _selected_cells(self) -> set[tuple[int, int]]:
        """tksheet seçimini Excel 1-tabanlı (satır, kolon) kümesine çevirir."""
        try:
            sel = self.sheet.get_selected_cells(get_rows=True, get_columns=True)
        except TypeError:
            sel = self.sheet.get_selected_cells()
        return {(r + 1, c + 1) for r, c in sel}

    def _require_file(self) -> bool:
        if not self.current_sheet:
            messagebox.showinfo(APP_TITLE, "Önce bir Excel dosyası yükleyin.")
            return False
        return True

    def mark_selection(self, unmark=False):
        if not self._require_file():
            return
        cells = self._selected_cells()
        if not cells:
            messagebox.showinfo(APP_TITLE, "Tabloda önce hücre, satır ya da kolon seçin.")
            return
        target = self.marks.setdefault(self.current_sheet, set())
        if unmark:
            target -= cells
        else:
            target |= cells
        self._redraw_marks()

    def _on_double_click(self, event):
        if not self.current_sheet:
            return
        try:
            r = self.sheet.identify_row(event, allow_end=False)
            c = self.sheet.identify_column(event, allow_end=False)
        except Exception:  # noqa: BLE001
            return
        if r is None or c is None:
            return
        target = self.marks.setdefault(self.current_sheet, set())
        cell = (r + 1, c + 1)
        target.discard(cell) if cell in target else target.add(cell)
        self._redraw_marks()

    def _selected_columns(self) -> set[int]:
        cols = {c + 1 for c in self.sheet.get_selected_columns()}
        if not cols:
            cols = {c for _, c in self._selected_cells()}
        return cols

    def _selected_rows(self) -> set[int]:
        rows = {r + 1 for r in self.sheet.get_selected_rows()}
        if not rows:
            rows = {r for r, _ in self._selected_cells()}
        return rows

    def mark_full_columns(self):
        """Seçili kolonları — ekranda görünmeyen satırlar dahil — baştan sona işaretler."""
        if not self._require_file():
            return
        cols = self._selected_columns()
        if not cols:
            messagebox.showinfo(APP_TITLE, "Önce bir kolon (ya da o kolondan bir hücre) seçin.")
            return
        self._add_columns(cols, skip_header=self._ask_skip_header())

    def _add_columns(self, cols, skip_header=False):
        total_rows, _ = self.sheet_dims[self.current_sheet]
        start = 2 if skip_header else 1
        target = self.marks.setdefault(self.current_sheet, set())
        for c in cols:
            target |= {(r, c) for r in range(start, total_rows + 1)}
        self._redraw_marks()

    def _ask_skip_header(self) -> bool:
        return messagebox.askyesno(
            APP_TITLE,
            "İlk satır başlık satırı mı?\n\n"
            "Evet → başlıklar okunur kalır, sadece veriler şifrelenir.\n"
            "Hayır → başlık da dahil kolonun tamamı şifrelenir.",
        )

    def mark_full_rows(self):
        if not self._require_file():
            return
        rows = self._selected_rows()
        if not rows:
            messagebox.showinfo(APP_TITLE, "Önce bir satır (ya da o satırdan bir hücre) seçin.")
            return
        _, total_cols = self.sheet_dims[self.current_sheet]
        target = self.marks.setdefault(self.current_sheet, set())
        for r in rows:
            target |= {(r, c) for c in range(1, total_cols + 1)}
        self._redraw_marks()

    def suggest_sensitive(self):
        """İlk satırdaki başlıklara bakıp hassas görünen kolonları işaretler."""
        if not self._require_file():
            return
        ws = self.wb[self.current_sheet]
        _, total_cols = self.sheet_dims[self.current_sheet]
        hits, names = [], []
        for c in range(1, total_cols + 1):
            header = ws.cell(row=1, column=c).value
            if header and SENSITIVE_RE.search(str(header)):
                hits.append(c)
                names.append(str(header).strip())
        if not hits:
            messagebox.showinfo(APP_TITLE, "Başlıklarda hassas görünen kolon bulunamadı.")
            return
        if messagebox.askyesno(
            APP_TITLE,
            "Şu kolonlar hassas görünüyor:\n\n  • " + "\n  • ".join(names)
            + "\n\nBaşlık satırı hariç tamamını işaretleyeyim mi?",
        ):
            self._add_columns(hits, skip_header=True)

    def clear_sheet_marks(self):
        if not self._require_file():
            return
        self.marks[self.current_sheet] = set()
        self._redraw_marks()

    def clear_all_marks(self):
        self.marks = {n: set() for n in self.sheet_names}
        self._redraw_marks()

    def _redraw_marks(self):
        if not self.current_sheet:
            return
        self.sheet.dehighlight_all()

        enc = self.enc_cells.get(self.current_sheet, set())
        visible_enc = [(r - 1, c - 1) for (r, c) in enc
                       if r <= MAX_DISPLAY_ROWS and c <= MAX_DISPLAY_COLS]
        if visible_enc:
            self.sheet.highlight_cells(cells=visible_enc, bg=ENC_BG, fg=ENC_FG, redraw=False)

        marked = self.marks.get(self.current_sheet, set())
        visible = [(r - 1, c - 1) for (r, c) in marked
                   if r <= MAX_DISPLAY_ROWS and c <= MAX_DISPLAY_COLS]
        if visible:
            self.sheet.highlight_cells(cells=visible, bg=MARK_BG, fg=MARK_FG, redraw=False)
        self.sheet.redraw()

        total = sum(len(v) for v in self.marks.values())
        here = len(marked)
        self.count_var.set(f"işaretli hücre: {here:,}  (tüm sayfalar: {total:,})".replace(",", "."))

    # ------------------------------------------------------------------ #
    # Şifrele / Çöz
    # ------------------------------------------------------------------ #
    def _key_for_encrypt(self, method) -> str | None:
        key = self.key_var.get()
        if not key:
            messagebox.showwarning(APP_TITLE, "Bir anahtar (parola) girin.")
            return None
        if key != self.key2_var.get():
            messagebox.showwarning(APP_TITLE, "İki anahtar alanı birbirini tutmuyor.")
            return None
        problem = ce.password_problem(key)
        if problem:
            messagebox.showwarning(APP_TITLE, problem)
            return None
        return key

    def on_encrypt(self):
        if self.busy or not self._require_file():
            return
        selections = {k: v for k, v in self.marks.items() if v}
        if not selections:
            messagebox.showinfo(APP_TITLE, "Hiç hücre işaretlemediniz.")
            return
        if self.enc_cells:
            messagebox.showerror(APP_TITLE, "Bu dosya zaten şifrelenmiş. Önce şifresini çözün.")
            return

        method = ce.METHODS_BY_LABEL[self.method_combo.get()]
        key = self._key_for_encrypt(method)
        if key is None:
            return

        total = sum(len(v) for v in selections.values())
        dest = xio.unique_path(xio.output_path(self.src_path, "encrypted"))
        if not messagebox.askokcancel(
            APP_TITLE,
            f"{total:,} hücre şifrelenecek.\n\n".replace(",", ".")
            + f"Yöntem: {method.label}\nÇıktı: {os.path.basename(dest)}\n"
            f"Klasör: {os.path.dirname(dest)}\n\n"
            "⚠ Anahtarı kaybederseniz veriler geri getirilemez.",
        ):
            return

        self._run_job(
            lambda progress: xio.encrypt_workbook(
                self.src_path, selections, method.id, key, dest_path=dest, progress=progress
            ),
            "Şifreleniyor…",
            lambda res: self._done("şifrelendi", res),
        )

    def on_decrypt(self):
        if self.busy or not self._require_file():
            return
        if not self.enc_cells:
            messagebox.showinfo(APP_TITLE, "Bu dosyada şifreli hücre yok.")
            return
        key = self.dkey_var.get()
        if not key:
            messagebox.showwarning(APP_TITLE, "Şifreyi çözmek için anahtarı girin.")
            return

        self._run_job(
            lambda progress: xio.decrypt_workbook(self.src_path, key, progress=progress),
            "Çözülüyor…",
            lambda res: self._done("çözüldü", res),
        )

    def _run_job(self, job, status, on_ok):
        """Ağır işi arka planda çalıştırır; arayüz donmaz."""
        self.busy = True
        self.enc_btn.state(["disabled"])
        self.dec_btn.state(["disabled"])
        self.progress.configure(value=0, maximum=100)
        self._set_status(status)

        def progress(done, total):
            self.after(0, lambda: self.progress.configure(value=done * 100 / max(total, 1)))

        def worker():
            try:
                result = job(progress)
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda: self._fail(exc))
            else:
                self.after(0, lambda: on_ok(result))

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self):
        self.busy = False
        self.enc_btn.state(["!disabled"])
        self.dec_btn.state(["!disabled"])

    def _fail(self, exc):
        self._finish()
        self.progress.configure(value=0)
        self._set_status("İşlem başarısız.")
        messagebox.showerror(APP_TITLE, str(exc))

    def _done(self, verb, result):
        dest, count = result
        self._finish()
        self.progress.configure(value=100)
        self._set_status(f"{count:,} hücre {verb} → {dest}".replace(",", "."))
        note = ""
        if verb == "şifrelendi":
            note = ("\n\n⚠ Kaynak dosya hâlâ ŞİFRESİZ duruyor:\n"
                    f"{self.src_path}\n"
                    "Veriyi korumak istiyorsanız o dosyayı güvenli şekilde silin "
                    "veya taşıyın.")
        if messagebox.askyesno(
            APP_TITLE,
            f"{count:,} hücre {verb}.\n\n".replace(",", ".")
            + f"Kaydedildi:\n{dest}" + note + "\n\nKlasörü açayım mı?",
        ):
            self._reveal(dest)

    @staticmethod
    def _reveal(path):
        """Çıktı dosyasını dosya yöneticisinde gösterir.

        Kabuk KULLANILMAZ: dosya adı tırnak/`$(...)` içerebilir ve os.system
        ile çalıştırılsaydı komut enjeksiyonuna açık olurdu.
        """
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", "-R", path], check=False)
            elif os.name == "nt":
                os.startfile(os.path.dirname(path))  # noqa: S606
            else:
                subprocess.run(["xdg-open", os.path.dirname(path)], check=False)
        except (OSError, subprocess.SubprocessError):
            pass


if __name__ == "__main__":
    App().mainloop()
