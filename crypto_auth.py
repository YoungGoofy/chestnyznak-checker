#!/usr/bin/env python3
"""
Автоматическое получение JWT-токена Честного Знака через УКЭП.

Метод авторизации: JWT flow — два шага:
  1. GET /auth/key → получаем uuid + data (challenge)
  2. Подписываем data сертификатом УКЭП → POST /auth/simpleSignIn → JWT-токен

Подпись выполняется через КриптоПро CSP:
- Windows: COM-объекты CAdESCOM (КриптоПро ECP / CSP 5.x) или legacy CPCSPStore (CSP 4.x)
- Fallback: cryptcp CLI (Windows / Linux)
"""
from __future__ import annotations

import base64
import json
import os
import platform
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# ── Константы ──────────────────────────────────────────────────────────

AUTH_KEY_URL = "https://markirovka.crpt.ru/api/v3/auth/key"
AUTH_SIGN_URL = "https://markirovka.crpt.ru/api/v3/auth/simpleSignIn"

TIMEOUT = 15

# Логирование
_log_fn = None


def set_log_fn(fn) -> None:
    """Устанавливает функцию логирования (для GUI)."""
    global _log_fn
    _log_fn = fn


def _get_log_fn():
    """Возвращает текущую функцию логирования."""
    if _log_fn:
        return _log_fn
    return print


# ── HTTP-хелперы ──────────────────────────────────────────────────────

def _http_get_json(url: str) -> dict:
    """GET-запрос, возвращает JSON."""
    req = Request(url, headers={"Accept": "application/json"})
    with urlopen(req, timeout=TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _http_post_json(url: str, body: dict) -> tuple[int, dict | str]:
    """POST-запрос с JSON, возвращает (status_code, json_or_text)."""
    data = json.dumps(body).encode("utf-8")
    req = Request(url, data=data, headers={
        "Content-Type": "application/json",
        "Accept": "application/json",
    })
    try:
        with urlopen(req, timeout=TIMEOUT) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return resp.status, result
    except HTTPError as e:
        try:
            err = json.loads(e.read().decode("utf-8"))
        except Exception:
            err = str(e)
        return e.code, err


# ── Список сертификатов УКЭП ──────────────────────────────────────────

# Константы хранилищ КриптоПро
CAPICOM_MEMORY_STORE = 0
CAPICOM_LOCAL_MACHINE_STORE = 1
CAPICOM_CURRENT_USER_STORE = 2

CAPICOM_STORE_OPEN_READ_ONLY = 0
CAPICOM_STORE_OPEN_READ_WRITE = 1
CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = 2

# OID КриптоПро для ИНН в Subject
OID_INN = "1.2.643.3.131.1.1"


def list_certificates() -> list[dict]:
    """
    Возвращает список доступных сертификатов УКЭП.

    Каждый сертификат: {
        "thumbprint": "XX XX ...",  — отпечаток (SHA-1)
        "subject": "...",            — субъект сертификата
        "issuer": "...",             — издатель
        "not_before": "...",         — дата начала
        "not_after": "...",          — дата окончания
        "inn": "1234567890",         — ИНН (если есть)
    }

    Порядок поиска:
    1. CAPICOM.Store (Windows Certificate Store — перехватывается КриптоПро CSP 5.x/4.x)
    2. CPCSPStore (CSP 4.x — legacy)
    3. cryptcp CLI (fallback)
    """
    log_fn = _get_log_fn()

    # 1. Пробуем CAdESCOM (современный интерфейс КриптоПро)
    certs = _list_certs_cadescom()
    if certs is not None and len(certs) > 0:
        log_fn(f"📋 Найдено сертификатов (CAdESCOM): {len(certs)}")
        return certs

    # 2. Пробуем legacy CPCSPStore
    certs = _list_certs_legacy_com()
    if certs is not None and len(certs) > 0:
        log_fn(f"📋 Найдено сертификатов (CPCSPStore): {len(certs)}")
        return certs

    # 3. Fallback: cryptcp CLI
    certs = _list_certs_cryptcp()
    if len(certs) > 0:
        log_fn(f"📋 Найдено сертификатов (cryptcp): {len(certs)}")
    else:
        log_fn("⚠ Сертификаты не найдены ни через COM, ни через cryptcp")

    return certs


def _list_certs_cadescom() -> list[dict] | None:
    """Список сертификатов через CAPICOM.Store (Windows Certificate Store, перехватывается КриптоПро CSP)."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client
    except ImportError:
        _get_log_fn()("⚠ pywin32 не установлен — COM-доступ к сертификатам недоступен. Установите: pip install pywin32")
        return None

    log_fn = _get_log_fn()
    certs = []

    # Пробуем хранилище "My" (личные сертификаты пользователя)
    try:
        log_fn("🔍 Открываю хранилище сертификатов (CAPICOM.Store)...")
        store = win32com.client.Dispatch("CAPICOM.Store")
        # Open(StoreLocation, StoreName, OpenMode)
        # CAPICOM_CURRENT_USER_STORE=2, CAPICOM_STORE_OPEN_READ_ONLY=0
        store.Open(
            CAPICOM_CURRENT_USER_STORE,
            "My",
            CAPICOM_STORE_OPEN_READ_ONLY,
        )
        log_fn(f"📂 Хранилище открыто, сертификатов: {store.Certificates.Count()}")

        for i in range(1, store.Certificates.Count() + 1):
            cert = store.Certificates.Item(i)
            info = _parse_cert_cadescom(cert)
            if info:
                certs.append(info)

        store.Close()
    except Exception as e:
        log_fn(f"⚠ CAPICOM.Store: {e}")
        # Пробуем альтернативный метод перебора
        try:
            store = win32com.client.Dispatch("CAPICOM.Store")
            store.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)
            for cert in store.Certificates:
                info = _parse_cert_cadescom(cert)
                if info:
                    certs.append(info)
            store.Close()
            log_fn(f"📂 CAPICOM.Store (MAX_ALLOWED): найдено {len(certs)} сертификатов")
        except Exception as e2:
            log_fn(f"⚠ CAPICOM.Store (альтернативный метод): {e2}")

    return certs if certs else None


def _parse_cert_cadescom(cert) -> dict | None:
    """Парсит сертификат из CAdESCOM."""
    try:
        subject = ""
        issuer = ""
        thumbprint = ""
        not_before = ""
        not_after = ""

        # Пробуем разные свойства (CAdESCOM vs CAPICOM)
        for attr in ["SubjectName", "Subject", "GetInfo(0)"]:
            try:
                if attr == "SubjectName":
                    subject = cert.SubjectName
                elif attr == "Subject":
                    subject = cert.Subject
                    break
            except Exception:
                continue

        for attr in ["IssuerName", "Issuer"]:
            try:
                if attr == "IssuerName":
                    issuer = cert.IssuerName
                elif attr == "Issuer":
                    issuer = cert.Issuer
                    break
            except Exception:
                continue

        try:
            thumbprint = cert.Thumbprint or ""
        except Exception:
            pass

        try:
            not_before = str(cert.ValidFromDate or "")
        except Exception:
            pass

        try:
            not_after = str(cert.ValidToDate or "")
        except Exception:
            pass

        # Извлекаем ИНН
        inn = _extract_inn_from_cert(cert)

        # Фильтруем: пропускаем корневые сертификаты без ИНН (CA-сертификаты)
        # УКЭП должен иметь ИНН или быть в хранилище "My"
        return {
            "thumbprint": thumbprint,
            "subject": subject,
            "issuer": issuer,
            "not_before": not_before,
            "not_after": not_after,
            "inn": inn,
        }
    except Exception:
        return None


def _extract_inn_from_cert(cert) -> str:
    """
    Извлекает ИНН из сертификата УКЭП.
    Пробует: ExtendedKeyUsage/OID, SubjectName, Subject (DN).
    """
    # Метод 1: OID КриптоПро для ИНН (1.2.643.3.131.1.1)
    try:
        # CAdESCOM: GetExtension/OID
        for method in ["GetExtension", "ExtendedKeyUsage"]:
            try:
                ext = getattr(cert, method)(OID_INN)
                if ext:
                    val = str(ext).strip()
                    if val.isdigit() and len(val) in (10, 12):
                        return val
            except Exception:
                continue
    except Exception:
        pass

    # Метод 2: Из SubjectName как строки
    try:
        subject = cert.SubjectName or cert.Subject or ""
        inn = _extract_inn_from_subject(subject)
        if inn:
            return inn
    except Exception:
        pass

    # Метод 3: Из Subject (DN-формат)
    try:
        subject = cert.Subject or ""
        inn = _extract_inn_from_subject(subject)
        if inn:
            return inn
    except Exception:
        pass

    return ""


def _list_certs_legacy_com() -> list[dict] | None:
    """Список сертификатов через legacy CPCSPStore.Store (КриптоПро CSP 4.x)."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client
    except ImportError:
        return None

    try:
        store = win32com.client.Dispatch("CPCSPStore.Store")
        store.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY)
        certs = []
        for cert in store.Certificates:
            info = _parse_cert_com_legacy(cert)
            if info:
                certs.append(info)
        store.Close()
        return certs
    except Exception as e:
        _get_log_fn()(f"⚠ CPCSPStore: {e}")
        return None


def _parse_cert_com_legacy(cert) -> dict | None:
    """Парсит COM-объект сертификата (legacy CPCSPStore)."""
    try:
        subject = cert.SubjectName or ""
        inn = _extract_inn_from_subject(subject)
        if not inn:
            inn = _extract_inn_from_cert(cert)
        return {
            "thumbprint": cert.Thumbprint or "",
            "subject": subject,
            "issuer": cert.IssuerName or "",
            "not_before": str(cert.ValidFromDate or ""),
            "not_after": str(cert.ValidToDate or ""),
            "inn": inn,
            "_com_cert": cert,
        }
    except Exception:
        return None


def _list_certs_cryptcp() -> list[dict]:
    """Список сертификатов через cryptcp (CLI)."""
    certs = []

    # Пробуем разные хранилища
    for store_flag in ["-u my", "-u root"]:
        try:
            parts = ["cryptcp", "-certstore"] + store_flag.split()
            result = subprocess.run(
                parts,
                capture_output=True, text=True, timeout=10,
                encoding="cp866" if platform.system() == "Windows" else "utf-8",
            )
            output = result.stdout + result.stderr

            current: dict = {}
            for line in output.splitlines():
                line = line.strip()
                if line.startswith("CertNum:"):
                    if current.get("thumbprint"):
                        certs.append(current)
                    current = {}
                elif line.startswith("Subject:"):
                    current["subject"] = line.split(":", 1)[1].strip()
                    current["inn"] = _extract_inn_from_subject(current.get("subject", ""))
                elif line.startswith("Issuer:"):
                    current["issuer"] = line.split(":", 1)[1].strip()
                elif "SHA1 Hash" in line or "SHA1" in line:
                    hash_part = line.split(":", 1)[1].strip() if ":" in line else ""
                    current["thumbprint"] = hash_part
            if current.get("thumbprint"):
                certs.append(current)
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue

    # Дедупликация по thumbprint
    seen = set()
    unique = []
    for c in certs:
        tp = c.get("thumbprint", "")
        if tp not in seen:
            seen.add(tp)
            unique.append(c)

    return unique


def _extract_inn_from_subject(subject: str) -> str:
    """Извлекает ИНН из Subject сертификата."""
    # Форматы: "INN=1234567890", "ИНН=1234567890", OID "1.2.643.3.131.1.1=#..."
    for sep in [", ", "; ", "\n"]:
        parts = subject.split(sep)
        for part in parts:
            part = part.strip()
            # INN=... / ИНН=...
            if part.upper().startswith("INN=") or part.upper().startswith("ИНН="):
                val = part.split("=", 1)[1].strip()
                if val.isdigit() and len(val) in (10, 12):
                    return val
            # OID КриптоПро: 1.2.643.3.131.1.1=#1604...
            if part.startswith(OID_INN + "="):
                val = part.split("=", 1)[1].strip()
                # Значение может быть в формате #hex или просто цифры
                if val.startswith("#"):
                    # hex-encoded, пытаемся декодировать
                    try:
                        hex_str = val[1:]
                        # Пропускаем ASN.1 tag + length (обычно #1604 + цифры ИНН в hex)
                        # Формат: #160431323334353637383930
                        # 16 = OID, 04 = OCTET STRING, потом hex цифр ИНН
                        if len(hex_str) >= 8:
                            inn_hex = hex_str[4:]  # пропускаем tag+length
                            inn_val = bytes.fromhex(inn_hex).decode("ascii", errors="ignore").strip()
                            if inn_val.isdigit() and len(inn_val) in (10, 12):
                                return inn_val
                    except Exception:
                        pass
                elif val.isdigit() and len(val) in (10, 12):
                    return val
    return ""


# ── Подпись данных ─────────────────────────────────────────────────────

def sign_data(data: bytes, thumbprint: str = "") -> bytes | None:
    """
    Подписывает данные УКЭП (attached CMS signature, DER).

    Возвращает CMS signature в DER (бинарный) или None при ошибке.

    Пробует методы в порядке:
    1. КриптоПро CAdESCOM COM (Windows, CSP 5.x)
    2. КриптоПро legacy COM (CPCSPStore, CSP 4.x)
    3. cryptcp CLI
    """
    log_fn = _get_log_fn()

    result = _sign_cadescom(data, thumbprint)
    if result is not None:
        return result

    result = _sign_legacy_com(data, thumbprint)
    if result is not None:
        return result

    result = _sign_cryptcp(data, thumbprint)
    if result is not None:
        return result

    log_fn("❌ Ни один метод подписи не сработал (CAdESCOM, CPCSPStore, cryptcp)")
    return None


def _find_cert_in_store(store, thumbprint: str):
    """Находит сертификат по thumbprint в COM Store."""
    for cert in store.Certificates:
        try:
            tp = cert.Thumbprint or ""
            if tp and tp.lower().replace(" ", "") == thumbprint.lower().replace(" ", ""):
                return cert
        except Exception:
            continue
    return None


def _sign_cadescom(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через CAdESCOM (КриптоПро CSP 5.x). Хранилище сертификатов: CAPICOM.Store."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client
    except ImportError:
        _get_log_fn()("⚠ pywin32 не установлен — COM-подпись недоступна. Установите: pip install pywin32")
        return None

    log_fn = _get_log_fn()

    try:
        # Открываем хранилище
        log_fn("🔍 Открываю хранилище сертификатов для подписи...")
        store = win32com.client.Dispatch("CAPICOM.Store")
        store.Open(
            CAPICOM_CURRENT_USER_STORE,
            "My",
            CAPICOM_STORE_OPEN_READ_ONLY,
        )

        cert_count = store.Certificates.Count
        log_fn(f"📂 В хранилище {cert_count} сертификатов")

        # Находим сертификат
        cert = None
        if thumbprint:
            cert = _find_cert_in_store(store, thumbprint)
            if cert:
                log_fn(f"✅ Сертификат найден по thumbprint")
            else:
                log_fn(f"⚠ Сертификат с thumbprint не найден, пробуем первый...")
        elif cert_count > 0:
            cert = store.Certificates.Item(1)

        store.Close()

        if not cert:
            log_fn("⚠ Сертификат не найден в хранилище My (CAdESCOM)")
            return None

        # Создаём Signer
        signer = win32com.client.Dispatch("CAdESCOM.CPSigner")
        signer.Certificate = cert

        # Создаём SignedData
        signed_data = win32com.client.Dispatch("CAdESCOM.CadesSignedData")
        # Content = base64(data)
        signed_data.Content = base64.b64encode(data).decode("ascii")

        # Подписываем (detached=False = attached)
        # CAdESCOM: Sign(Signer, Detached, EncodingType)
        # EncodingType: 0 = Base64, 1 = XML
        signature_b64 = signed_data.Sign(signer, False, 0)

        # Декодируем из base64 → DER
        return base64.b64decode(signature_b64)

    except Exception as e:
        log_fn(f"⚠ CAdESCOM подпись: {e}")
        return None


def _sign_legacy_com(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через legacy CPCSPStore (КриптоПро CSP 4.x)."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client
    except ImportError:
        return None

    log_fn = _get_log_fn()

    try:
        store = win32com.client.Dispatch("CPCSPStore.Store")
        store.Open()

        cert = None
        if thumbprint:
            cert = _find_cert_in_store(store, thumbprint)
        elif store.Certificates.Count > 0:
            cert = store.Certificates.Item(1)

        if not cert:
            store.Close()
            log_fn("⚠ Сертификат не найден (CPCSPStore)")
            return None

        signer = win32com.client.Dispatch("CPSigner.Signer")
        signer.Certificate = cert

        signed_data = win32com.client.Dispatch("CPSignedData.SignedData")
        signed_data.Content = base64.b64encode(data).decode("ascii")

        # Исправленный баг: было signedData (опечатка), сейчас signed_data
        signature_b64 = signed_data.SignCPS(signer, True, 0)

        store.Close()
        return base64.b64decode(signature_b64)

    except Exception as e:
        log_fn(f"⚠ CPCSPStore подпись: {e}")
        try:
            store.Close()
        except Exception:
            pass
        return None


def _sign_cryptcp(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через cryptcp CLI (Windows/Linux)."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as f_in:
        f_in.write(data)
        in_path = f_in.name

    out_path = in_path + ".sig"

    try:
        cmd = ["cryptcp", "-sign", "-der"]
        if thumbprint:
            thumb_clean = thumbprint.lower().replace(" ", "")
            cmd += ["-cert", thumb_clean]
        cmd += ["-u", "my", in_path, out_path]

        result = subprocess.run(
            cmd,
            capture_output=True, timeout=30,
            encoding="cp866" if platform.system() == "Windows" else "utf-8",
        )

        if result.returncode == 0 and Path(out_path).exists():
            return Path(out_path).read_bytes()
        return None
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None
    finally:
        for p in (in_path, out_path):
            try:
                Path(p).unlink()
            except OSError:
                pass


# ── Авторизация ────────────────────────────────────────────────────────

def auth_jwt(thumbprint: str = "") -> tuple[bool, str]:
    """
    Получает JWT-токен через двухшаговый flow.

    Шаг 1: GET /auth/key → Получаем uuid + data (challenge)
    Шаг 2: Подписываем data → POST /auth/simpleSignIn

    Возвращает (success: bool, token_or_error: str).
    """
    log_fn = _get_log_fn()

    # Шаг 1: Получаем challenge
    log_fn("📡 Получаю challenge с сервера...")
    try:
        challenge = _http_get_json(AUTH_KEY_URL)
    except Exception as e:
        return False, f"Не удалось получить challenge: {e}"

    uuid = challenge.get("uuid", "")
    data_str = challenge.get("data", "")
    if not uuid or not data_str:
        return False, "Сервер вернул некорректный challenge."

    log_fn("🔐 Подписываю challenge сертификатом УКЭП...")

    # Шаг 2: Подписываем challenge
    signature = sign_data(data_str.encode("utf-8"), thumbprint)
    if signature is None:
        return False, (
            "Не удалось подписать данные. Убедитесь, что:\n"
            "• КриптоПро CSP установлен\n"
            "• УКЭП (RuToken/USB-токен) подключён\n"
            "• Сертификат доступен в хранилище\n"
            "• pywin32 установлен (для COM-доступа)"
        )

    sig_b64 = base64.b64encode(signature).decode("ascii")

    # Отправляем подпись
    log_fn("📡 Отправляю подпись на сервер...")
    status, response = _http_post_json(AUTH_SIGN_URL, {
        "uuid": uuid,
        "data": sig_b64,
    })

    if status == 200 and isinstance(response, dict):
        token = response.get("token", "")
        if token:
            log_fn("✅ JWT-токен получен!")
            return True, token

    error_msg = _parse_auth_error(status, response)
    log_fn(f"❌ {error_msg}")
    return False, error_msg


def _parse_auth_error(status: int, response) -> str:
    """Расшифровка ошибок авторизации."""
    if isinstance(response, dict):
        msg = response.get("error_message", "") or response.get("message", "")
        if msg:
            if status == 400:
                return f"Ошибка запроса (HTTP 400): {msg}"
            if status == 403:
                return f"Доступ запрещён (HTTP 403): {msg}"
    if isinstance(response, str):
        if status == 400:
            return f"Ошибка запроса (HTTP 400): {response}"
        if status == 403:
            return f"Доступ запрещён (HTTP 403): {response}"
    if status == 403:
        return "Доступ запрещён (HTTP 403). Пользователь не найден или неактивен."
    return f"Ошибка авторизации (HTTP {status})."