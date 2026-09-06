# Excel Encryptor

Excel dosyalarındaki gizli verileri **hücre bazında** şifreleyen masaüstü programı.
Dosyanın yapısı, formülleri ve şifrelenmemiş hücreleri olduğu gibi kalır — sadece
seçtiğiniz hücreler okunamaz hale gelir.

> Cell-level encryption for Excel files. Pick the cells that hold sensitive data,
> choose a key, and get an encrypted copy alongside the original. macOS app and
> Windows executable, no installation required.

![platform](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-lightgrey)
![license](https://img.shields.io/badge/license-MIT-blue)

## İndirme

En son sürümü [Releases](../../releases/latest) sayfasından indirin:

| Dosya | Platform | Nasıl çalıştırılır |
|---|---|---|
| `ExcelEncryptor.exe` | Windows 10/11 (64-bit) | Çift tıklayın. Kurulum ve Python gerekmez. |
| `Excel Encryptor-macos.zip` | macOS 12+ (Apple Silicon) | Açın, `.app`'i Uygulamalar'a sürükleyin, çift tıklayın. |

İlk açılışta işletim sistemi uyarır (dosyalar kod imzalı değil):

- **Windows:** SmartScreen → *Daha fazla bilgi* → *Yine de çalıştır*
- **macOS:** uygulamaya sağ tık → *Aç* → açılan pencerede yine *Aç*

İndirdiğinizi doğrulamak isterseniz her sürümde `SHA256SUMS.txt` yayınlanıyor.

## Kullanım

### Şifreleme

1. **Excel Yükle** — `.xlsx` / `.xlsm` seçin. Çok sayfalı dosyalarda sağ üstten sayfa değiştirin.
2. **Gizlenecek hücreleri işaretleyin** (işaretliler turuncu görünür):
   - Tabloda hücre/aralık seçip **Seçileni İşaretle**
   - Kolon başlığına tıklayıp **Kolonun Tamamını İşaretle** — ekranda görünmeyen satırlar dahil; başlık satırını hariç tutmayı sorar
   - Satır numarasına tıklayıp **Satırın Tamamını İşaretle**
   - Tek hücreye **çift tıklamak** işareti açıp kapatır
   - **Hassas Kolonları Öner** — başlıklarda TC, IBAN, maaş, e-posta, şifre, reçete gibi kalıpları arar
3. **Yöntemi seçip anahtarı iki kez girin.**
4. **ŞİFRELE** → kaynak dosyayla aynı klasöre `Dosya_adı_encrypted.xlsx` yazılır.
   Orijinal dosyaya dokunulmaz.

### Şifre çözme

**Şifre Çöz** sekmesinde şifreli dosyayı yükleyin — program dosyayı tanıyıp otomatik
bu sekmeye geçer, şifreli hücreleri mavi gösterir. Anahtarı girin (yöntem dosyadan
okunur), **ŞİFREYİ ÇÖZ** → `Dosya_adı_decrypted.xlsx`.

Anahtar yanlışsa program işleme başlamadan uyarır.

## Şifreleme yöntemleri

Hepsi **kimlik doğrulamalı** (AEAD): şifreli veri değiştirilirse çözme sessizce
yanlış sonuç vermez, hata verir.

| Yöntem | Not |
|---|---|
| **AES-256-GCM** | Önerilen. Aynı değer her hücrede farklı şifrelenir. |
| **ChaCha20-Poly1305** | Eşdeğer güvenlik; donanım AES desteği olmayan makinelerde daha hızlı. |
| **Fernet** (AES-128-CBC + HMAC) | Standart zarf biçimi. |
| **AES-256-GCM — Deterministik** | Aynı değer → aynı şifre; gruplama korunur. ⚠ Tekrar eden değerleri ele verir. |

Anahtar **PBKDF2-HMAC-SHA256** ile 600.000 tur ve dosyaya özel rastgele tuz
kullanılarak 256-bit anahtara çevrilir. Anahtarın kendisi hiçbir yere yazılmaz.

Güvenlik incelemesinin tamamı, tehdit modeli ve bilinen sınırlar için
**[SECURITY.md](SECURITY.md)**.

## Dosya biçimi

Her şifreli hücre kendi kendini tanımlayan bir metne dönüşür:

```
ENC1:aesgcm:BASE64…
```

Hücrenin orijinal tipi (sayı / tarih / saat / mantıksal / metin) şifreli yükün
içinde saklanır, bu yüzden çözme sonrası sayılar yine sayı, tarihler yine tarih
olarak geri gelir.

Tuz, yöntem ve tur sayısı dosyadaki gizli `_ENCMETA_` sayfasında tutulur.
**Bu sayfa silinirse dosya çözülemez.**

## Uyarılar

- **Anahtarı kaybederseniz veri geri getirilemez.** Anahtarı şifreli dosyadan
  ayrı, güvenli bir yerde saklayın.
- **Kaynak dosya şifresiz kalır.** Program orijinali değiştirmez; veriyi
  korumak istiyorsanız kaynağı siz güvenli şekilde silin.
- Çıktı `openpyxl` ile yazılır: hücre değerleri, formüller, biçimler ve kolon
  genişlikleri korunur; ancak **grafikler, gömülü görseller, pivot tablolar ve
  `.xlsm` makroları aktarılmaz.** Önemliyse önce orijinali yedekleyin.
- Şifrelenen hücrede formül varsa formülün kendisi şifrelenir; o hücre artık
  hesaplanmaz.
- Ekranda en fazla 5000 satır × 200 kolon gösterilir; kolon/satır işaretlemesi
  dosyanın **tamamına** uygulanır.
- Aynı isimde çıktı varsa üzerine yazılmaz, sonuna `(1)`, `(2)` eklenir.

## Geliştirme

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python src/app.py          # çalıştır
.venv/bin/python -m pytest -q        # testler
```

**Gereksinim:** Python 3.9+ ve **Tk 8.6+**. macOS'ta Apple'ın kendi Python'ı
Tk 8.5 ile gelir ve tablo bileşenini çalıştıramaz —
[python.org](https://python.org/downloads) sürümünü ya da
`brew install python-tk@3.14` kullanın.

Paketleme (yerel):

```bash
bash packaging/make_icns.sh                                # yalnızca macOS
pyinstaller packaging/ExcelEncryptor.spec --noconfirm
```

Yayın: `v` ile başlayan bir etiket gönderin (`git tag v2.0 && git push --tags`);
CI her iki platformda test edip paketleri Releases'e yükler.

## Depo yapısı

| Yol | İçerik |
|---|---|
| `src/app.py` | Tkinter arayüzü: Şifrele / Şifre Çöz sekmeleri, hücre seçimi |
| `src/excel_io.py` | Excel okuma, şifreleyip/çözüp kaydetme, metadata |
| `src/crypto_engine.py` | Yöntem kataloğu, anahtar türetme, token biçimi |
| `tests/` | Şifreleme turu, tip korunumu ve güvenlik regresyon testleri |
| `packaging/` | PyInstaller yapılandırması ve ikon üretimi |
| `.github/workflows/` | Test + paketleme + CodeQL taraması |

## Lisans

MIT — bkz. [LICENSE](LICENSE).
