"""Şifreleme motoru — Excel hücre değerlerini metne çevirip geri alır.

Her şifreli hücre kendi kendini tanımlayan bir token'a dönüşür:

    ENC1:<yöntem-id>:<base64-yük>

Hücrenin orijinal tipi (sayı / tarih / bool / metin) yükün içinde saklanır,
böylece çözme sonrası sayılar yine sayı olarak geri gelir.

Güvenlik notu: burada yalnızca **kimlik doğrulamalı** (AEAD) şifreleme
yöntemleri bulunur. XOR, Vigenère ve Base64 gibi gerçek koruma sağlamayan
yöntemler, kullanıcıya yanlış güven verdikleri için bilinçli olarak
kaldırılmıştır (bkz. SECURITY.md).
"""

from __future__ import annotations

import base64
import datetime as _dt
import hmac
import hashlib
import json
import os

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives.ciphers.aead import AESGCM, ChaCha20Poly1305
from cryptography.exceptions import InvalidTag
from cryptography.fernet import Fernet, InvalidToken

TOKEN_PREFIX = "ENC1"

# OWASP 2023'ün PBKDF2-HMAC-SHA256 için önerdiği alt sınır.
KDF_ITERATIONS = 600_000
SALT_BYTES = 16
KEY_BYTES = 32

META_SHEET = "_ENCMETA_"
META_VERSION = 2

MIN_PASSWORD_LENGTH = 8


class CryptoError(Exception):
    """Anahtar hatalı ya da veri bozuk."""


# --------------------------------------------------------------------------- #
# Yöntem kataloğu — hepsi kimlik doğrulamalı (AEAD)
# --------------------------------------------------------------------------- #

class Method:
    def __init__(self, mid, label, note="", warning=""):
        self.id = mid
        self.label = label
        self.note = note
        self.warning = warning


METHODS = [
    Method("aesgcm", "AES-256-GCM  (önerilen)",
           note="Kimlik doğrulamalı. Aynı değer her hücrede farklı şifrelenir."),
    Method("chacha", "ChaCha20-Poly1305",
           note="AES-GCM ile eşdeğer güvenlik; donanım AES desteği olmayan "
                "makinelerde daha hızlı."),
    Method("fernet", "Fernet  (AES-128-CBC + HMAC)",
           note="Standart, kimlik doğrulamalı zarf biçimi."),
    Method("aesgcm-det", "AES-256-GCM — Deterministik",
           note="Aynı değer her yerde aynı şifreye dönüşür; gruplama ve "
                "eşleştirme korunur.",
           warning="Tekrar eden değerleri ele verir: saldırgan hangi hücrelerin "
                   "aynı veriyi taşıdığını görebilir. Yalnızca eşleştirme "
                   "gerektiğinde kullanın."),
]

METHODS_BY_ID = {m.id: m for m in METHODS}
METHODS_BY_LABEL = {m.label: m for m in METHODS}
DEFAULT_METHOD = "aesgcm"


# --------------------------------------------------------------------------- #
# Anahtar türetme
# --------------------------------------------------------------------------- #

def new_salt() -> bytes:
    return os.urandom(SALT_BYTES)


def derive_key(password: str, salt: bytes, iterations: int = KDF_ITERATIONS) -> bytes:
    """Parolayı PBKDF2-HMAC-SHA256 ile 32 baytlık anahtara çevirir."""
    if not password:
        raise CryptoError("Anahtar boş olamaz.")
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=KEY_BYTES,
                     salt=salt, iterations=iterations)
    return kdf.derive(password.encode("utf-8"))


def password_problem(password: str) -> str | None:
    """Parola yeterince güçlü mü? Sorun varsa açıklamasını döndürür."""
    if len(password) < MIN_PASSWORD_LENGTH:
        return f"Anahtar en az {MIN_PASSWORD_LENGTH} karakter olmalı."
    kinds = sum((
        any(c.islower() for c in password),
        any(c.isupper() for c in password),
        any(c.isdigit() for c in password),
        any(not c.isalnum() for c in password),
    ))
    if kinds < 2:
        return ("Anahtar en az iki farklı karakter türü içermeli "
                "(küçük/büyük harf, rakam, sembol).")
    return None


# --------------------------------------------------------------------------- #
# Hücre değeri <-> bayt
# --------------------------------------------------------------------------- #

def _pack(value) -> bytes:
    """Hücre değerini tipiyle birlikte JSON'a gömer."""
    if isinstance(value, bool):
        kind, payload = "b", value
    elif isinstance(value, int):
        kind, payload = "i", value
    elif isinstance(value, float):
        kind, payload = "f", value
    elif isinstance(value, _dt.datetime):
        kind, payload = "dt", value.isoformat()
    elif isinstance(value, _dt.date):
        kind, payload = "d", value.isoformat()
    elif isinstance(value, _dt.time):
        kind, payload = "t", value.isoformat()
    else:
        kind, payload = "s", str(value)
    return json.dumps([kind, payload], ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def _unpack(raw: bytes):
    try:
        kind, payload = json.loads(raw.decode("utf-8"))
    except Exception as exc:  # noqa: BLE001
        raise CryptoError("Çözülen veri okunamadı (anahtar hatalı olabilir).") from exc
    if kind == "b":
        return bool(payload)
    if kind == "i":
        return int(payload)
    if kind == "f":
        return float(payload)
    if kind == "dt":
        return _dt.datetime.fromisoformat(payload)
    if kind == "d":
        return _dt.date.fromisoformat(payload)
    if kind == "t":
        return _dt.time.fromisoformat(payload)
    return payload


# --------------------------------------------------------------------------- #
# Ham şifreleme katmanı (bayt -> bayt)
# --------------------------------------------------------------------------- #

def _encrypt_bytes(method_id: str, key: bytes, data: bytes) -> bytes:
    if method_id == "aesgcm":
        nonce = os.urandom(12)
        return nonce + AESGCM(key).encrypt(nonce, data, None)

    if method_id == "aesgcm-det":
        # SIV benzeri: nonce açık metinden türetilir, böylece aynı girdi aynı
        # çıktıyı verir. Nonce anahtara bağlı olduğu için tahmin edilemez.
        nonce = hmac.new(key, data, hashlib.sha256).digest()[:12]
        return nonce + AESGCM(key).encrypt(nonce, data, None)

    if method_id == "chacha":
        nonce = os.urandom(12)
        return nonce + ChaCha20Poly1305(key).encrypt(nonce, data, None)

    if method_id == "fernet":
        return Fernet(base64.urlsafe_b64encode(key)).encrypt(data)

    raise CryptoError(f"Bilinmeyen yöntem: {method_id}")


def _decrypt_bytes(method_id: str, key: bytes, blob: bytes) -> bytes:
    try:
        if method_id in ("aesgcm", "aesgcm-det"):
            return AESGCM(key).decrypt(blob[:12], blob[12:], None)
        if method_id == "chacha":
            return ChaCha20Poly1305(key).decrypt(blob[:12], blob[12:], None)
        if method_id == "fernet":
            return Fernet(base64.urlsafe_b64encode(key)).decrypt(blob)
    except (InvalidTag, InvalidToken, ValueError, TypeError) as exc:
        raise CryptoError("Hücre çözülemedi — anahtar hatalı ya da veri bozulmuş.") from exc

    raise CryptoError(f"Bilinmeyen yöntem: {method_id}")


# --------------------------------------------------------------------------- #
# Token seviyesi API — arayüz bunları kullanır
# --------------------------------------------------------------------------- #

def encrypt_value(value, method_id: str, key: bytes) -> str:
    blob = _encrypt_bytes(method_id, key, _pack(value))
    return f"{TOKEN_PREFIX}:{method_id}:{base64.b64encode(blob).decode()}"


def is_token(value) -> bool:
    return isinstance(value, str) and value.startswith(TOKEN_PREFIX + ":") and value.count(":") >= 2


def token_method(value: str) -> str:
    return value.split(":", 2)[1]


def decrypt_value(token: str, key: bytes):
    _, method_id, b64 = token.split(":", 2)
    if method_id not in METHODS_BY_ID:
        raise CryptoError(f"Dosyada tanınmayan şifreleme yöntemi: {method_id}")
    try:
        blob = base64.b64decode(b64, validate=True)
    except Exception as exc:  # noqa: BLE001
        raise CryptoError("Şifreli hücre bozuk (base64 çözülemedi).") from exc
    return _unpack(_decrypt_bytes(method_id, key, blob))
