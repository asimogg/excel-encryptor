# Güvenlik

Bu belge programın tehdit modelini, yapılan güvenlik incelemesinde bulunan ve
giderilen sorunları, ve bilinen sınırları anlatır.

## Tehdit modeli

**Neye karşı korur:** Şifreli dosyayı ele geçiren biri — e-postayla yanlış kişiye
giden, ortak sürücüde duran, çalınan diskteki dosya — anahtarı bilmeden
şifrelenmiş hücrelerin içeriğini okuyamaz ve fark edilmeden değiştiremez.

**Neye karşı korumaz:**

- **Zayıf anahtar.** Koruma tamamen anahtarın gücüne bağlıdır. Sözlükte geçen
  bir parola çevrimdışı denemeyle kırılabilir.
- **Ele geçirilmiş bilgisayar.** Keylogger, ekran kaydı ya da bellek okuma
  yapabilen bir saldırgana karşı hiçbir garanti yoktur.
- **Kaynak dosya.** Program orijinali değiştirmez; şifresiz kopya diskinizde
  kalmaya devam eder. Onu siz güvenli şekilde silmelisiniz.
- **Şifrelenmemiş hücreler.** Yalnızca işaretlediğiniz hücreler korunur;
  sayfa adları, kolon başlıkları ve diğer hücreler açık kalır.

## İncelemede bulunan ve giderilen sorunlar

| # | Sorun | Etki | Çözüm |
|---|---|---|---|
| 1 | **Komut enjeksiyonu** — çıktı klasörünü açan kod dosya yolunu `os.system()` ile kabuğa gönderiyordu | `ev"; rm -rf ~; echo ".xlsx` gibi bir dosya adı, kullanıcı "klasörü aç" dediğinde rastgele komut çalıştırabilirdi | Kabuk tamamen devre dışı: `subprocess.run([...])` liste biçiminde çağrılıyor. Regresyon testi: `test_reveal_does_not_pass_filenames_through_a_shell` |
| 2 | **Sahte güvenlik veren şifreleme yöntemleri** — XOR, Vigenère ve Base64 | Kullanıcı "şifreledim" sanırken veri pratikte korumasızdı; Base64 tek satır kodla geri çevrilir | Üç yöntem de kaldırıldı. Katalogda yalnızca kimlik doğrulamalı (AEAD) yöntemler var. Test: `test_insecure_methods_are_not_offered` |
| 3 | **Kimlik doğrulaması olmayan şifreleme** | Şifreli hücre değiştirildiğinde program sessizce yanlış veri döndürebilirdi | Kalan dört yöntemin hepsi AEAD. Test: `test_tampered_ciphertext_is_detected` |
| 4 | **Parola deneme oracle'ı** — dosyada `HMAC(anahtar, sabit)` doğrulama etiketi tutuluyordu | Saldırgan, şifreli veriye hiç dokunmadan parola denemelerini doğrulayabiliyordu | Etiket kaldırıldı; anahtar, ilk hücreyi çözmeyi deneyerek doğrulanıyor. Test: `test_metadata_leaks_nothing_beyond_what_decryption_needs` |
| 5 | **Hedef haritası sızıntısı** — metadata hangi hücrelerin gizlendiğini listeliyordu | Saldırgana "kıymetli veri şurada" yol haritası veriyordu | Harita kaldırıldı (zaten gereksizdi; şifreli hücreler token önekinden bulunuyor) |
| 6 | **Zayıf anahtar türetme** — PBKDF2 200.000 tur | Kaba kuvvet saldırısı OWASP eşiğinin altında maliyetliydi | 600.000 tura çıkarıldı (OWASP 2023 alt sınırı). Test: `test_kdf_meets_owasp_minimum` |
| 7 | **Parola politikası yok** — 4 karakter yeterliydi | Anlamsız derecede zayıf anahtarlar kabul ediliyordu | En az 8 karakter ve iki farklı karakter türü zorunlu |
| 8 | **XML saldırılarına açıklık** — `openpyxl` XXE ve "billion laughs" varlık genişletmesine karşı koruma sağlamaz | Dışarıdan gelen bir `.xlsx`, dosya okuma ya da bellek tüketimi saldırısı taşıyabilirdi | `defusedxml.defuse_stdlib()` açılışta çağrılıyor; ayrıca 5 milyon hücre üst sınırı. Test: `test_xml_parser_is_hardened` |
| 9 | **Yakalanmayan `InvalidTag`** | Bozuk dosyada ham istisna arayüze sızıyordu | `_decrypt_bytes` artık `InvalidTag`'i de yakalayıp anlaşılır hata veriyor |
| 10 | **Gatekeeper devre dışı bırakma** — kurulum betiği `xattr -dr com.apple.quarantine` çalıştırıyordu | macOS'un indirilen dosya kontrolünü tüm klasör için kapatıyordu | Kurulum betiği tamamen kaldırıldı; uygulama artık imzalı `.app` olarak paketleniyor |
| 11 | **Depoda gömülü ikili dosyalar** — 67 MB'lık CPython dağıtımı ve `.whl` dosyaları | Denetlenemeyen tedarik zinciri; güncellenmeyen bağımlılıklar | Hepsi kaldırıldı. Bağımlılıklar `requirements.txt`'te alt sınırlarla sabitlendi, CI'da `pip-audit` ile taranıyor |

## Uygulanan kriptografi

| Katman | Seçim |
|---|---|
| Anahtar türetme | PBKDF2-HMAC-SHA256, 600.000 tur, dosya başına 16 baytlık rastgele tuz (`os.urandom`) |
| Anahtar uzunluğu | 256 bit |
| Şifreleme | AES-256-GCM (varsayılan), ChaCha20-Poly1305, Fernet, AES-256-GCM deterministik |
| Nonce | Hücre başına 12 bayt rastgele; deterministik modda `HMAC(anahtar, açık_metin)[:12]` (SIV benzeri) |
| Kimlik doğrulama | Her yöntemde yerleşik (AEAD) |

Anahtarın kendisi hiçbir yere yazılmaz. Dosyada yalnızca tuz, yöntem adı ve
tur sayısı saklanır.

### Deterministik mod hakkında

`AES-256-GCM — Deterministik` seçildiğinde aynı değer her yerde aynı şifreye
dönüşür. Bu, gruplama/eşleştirme gerektiren analizler için kullanışlıdır ama
**tekrar eden değerleri ele verir**: saldırgan hangi satırların aynı veriyi
taşıdığını görebilir. Arayüz bu uyarıyı yöntem seçilince gösterir.

## Bilinen sınırlar

- **Anahtar bellekte açık durur.** Python'da metin nesneleri değiştirilemez
  olduğu için anahtarı bellekten güvenilir şekilde silmek mümkün değil.
- **Şifrelenen hücre sayısı ve konumu gizlenmez.** Şifreli hücreler dosyada
  görünür; hangi kolonların hassas kabul edildiği anlaşılabilir.
- **Uygulamalar imzalı değil.** Apple Developer / Windows kod imzalama
  sertifikası olmadığı için ilk açılışta Gatekeeper ve SmartScreen uyarır.
  Yayınlanan dosyaların SHA-256 özetleri her sürümde `SHA256SUMS.txt`
  içinde verilir; indirdiğinizi doğrulayabilirsiniz.
- **`.xlsm` makroları aktarılmaz.** `openpyxl` VBA kodunu korumaz — makrolar
  şifreli kopyada bulunmaz. (Güvenlik açısından bu istenen davranıştır.)

## Açık bildirimi

Bir güvenlik açığı bulursanız lütfen herkese açık issue açmadan önce
depo sahibiyle iletişime geçin.
