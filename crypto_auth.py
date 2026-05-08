#!/usr/bin/env python3
"""
Автоматическое получение JWT-токена Честного Знака через УКЭП.

Метод авторизации: JWT flow — два шага:
  1. GET /auth/key → получаем uuid + data (challenge)
  2. Подписываем data сертификатом УКЭП → POST /auth/simpleSignIn → JWT-токен

Подпись выполняется через КриптоПро CSP:
- Windows: COM-объекты через win32com (если pywin32 установлен) или cryptcp.exe
- Linux: cryptcp / csptest
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
    1. Windows COM (win32com + КриптоПро CSP)
    2. cryptcp CLI (Windows / Linux)
    """
    certs = _list_certs_com()
    if certs is not None:
        return certs
    return _list_certs_cryptcp()


def _list_certs_com() -> list[dict] | None:
    """Список сертификатов через КриптоПро COM (Windows)."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client
    except ImportError:
        return None

    try:
        store = win32com.client.Dispatch("CPCSPStore.Store")
        store.Open()
        certs = []
        for cert in store.Certificates:
            info = _parse_cert_com(cert)
            if info:
                certs.append(info)
        store.Close()
        return certs
    except Exception:
        return None


def _parse_cert_com(cert) -> dict | None:
    """Парсит COM-объект сертификата."""
    try:
        subject = cert.SubjectName
        # Извлекаем ИНН из субъекта
        inn = _extract_inn_from_subject(subject)
        return {
            "thumbprint": cert.Thumbprint or "",
            "subject": subject,
            "issuer": cert.IssuerName or "",
            "not_before": str(cert.ValidFromDate or ""),
            "not_after": str(cert.ValidToDate or ""),
            "inn": inn,
            # Сохраняем ссылку на COM-объект для подписи
            "_com_cert": cert,
        }
    except Exception:
        return None


def _list_certs_cryptcp() -> list[dict]:
    """Список сертификатов через cryptcp (CLI)."""
    certs = []
    # cryptcp -certstore -u my  — список сертификатов пользователя
    try:
        result = subprocess.run(
            ["cryptcp", "-certstore", "-u", "my"],
            capture_output=True, text=True, timeout=10,
            encoding="cp866" if platform.system() == "Windows" else "utf-8",
        )
        output = result.stdout + result.stderr
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return certs

    # Парсим вывод cryptcp
    # Формат:
    #   CertNum: 0
    #   Subject: ...
    #   Issuer: ...
    #   SHA1 Hash: XX XX ...
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
            # "SHA1 Hash: XX XX XX ..."
            hash_part = line.split(":", 1)[1].strip() if ":" in line else ""
            current["thumbprint"] = hash_part
    if current.get("thumbprint"):
        certs.append(current)

    return certs


def _extract_inn_from_subject(subject: str) -> str:
    """Извлекает ИНН из Subject сертификата."""
    # Форматы: "INN=1234567890" или "ИНН=1234567890" или "2.5.4.16=#1234567890"
    for sep in [", ", "; "]:
        parts = subject.split(sep)
        for part in parts:
            part = part.strip()
            if part.upper().startswith("INN=") or part.upper().startswith("ИНН="):
                val = part.split("=", 1)[1].strip()
                if val.isdigit() and len(val) in (10, 12):
                    return val
    return ""


# ── Подпись данных ─────────────────────────────────────────────────────

def sign_data(data: bytes, thumbprint: str = "") -> bytes | None:
    """
    Подписывает данные УКЭП (attached CMS signature).

    Возвращает CMS signature в DER (бинарный) или None при ошибке.

    Пробует методы в порядке:
    1. КриптоПро COM (Windows)
    2. cryptcp CLI
    """
    result = _sign_com(data, thumbprint)
    if result is not None:
        return result
    return _sign_cryptcp(data, thumbprint)


def _sign_com(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через КриптоПро COM (Windows)."""
    if platform.system() != "Windows":
        return None
    try:
        import win32com.client
    except ImportError:
        return None

    try:
        signer = win32com.client.Dispatch("CPSigner.Signer")
        if thumbprint:
            # Находим сертификат по thumbprint
            store = win32com.client.Dispatch("CPCSPStore.Store")
            store.Open()
            for cert in store.Certificates:
                if cert.Thumbprint and cert.Thumbprint.lower().replace(" ", "") == thumbprint.lower().replace(" ", ""):
                    signer.Certificate = cert
                    break
            store.Close()

        # Создаём объект для подписи
        signed_data = win32com.client.Dispatch("CPSignedData.SignedData")
        signed_data.Content = base64.b64encode(data).decode("ascii")

        # Подписываем (attached = True)
        signature = signedData.SignCPS(signer, True, 0)

        # Декодируем из base64
        return base64.b64decode(signature)
    except Exception:
        return None


def _sign_cryptcp(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через cryptcp CLI (Windows/Linux)."""
    # Создаём временные файлы
    with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as f_in:
        f_in.write(data)
        in_path = f_in.name

    out_path = in_path + ".sig"

    try:
        cmd = ["cryptcp", "-sign", "-der"]
        if thumbprint:
            # cryptcp -sign -cert <thumbprint> -der
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
        # Удаляем временные файлы
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

    log_fn(f"🔐 Подписываю challenge сертификатом УКЭП...")

    # Шаг 2: Подписываем challenge
    signature = sign_data(data_str.encode("utf-8"), thumbprint)
    if signature is None:
        return False, (
            "Не удалось подписать данные. Убедитесь, что:\n"
            "• КриптоПро CSP установлен\n"
            "• УКЭП (RuToken/USB-токен) подключён\n"
            "• Сертификат доступен в хранилище"
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
            log_fn(f"✅ JWT-токен получен!")
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


# ── Логирование ────────────────────────────────────────────────────────

_log_fn = None


def set_log_fn(fn) -> None:
    """Устанавливает функцию логирования (для GUI)."""
    global _log_fn
    _log_fn = fn


def _get_log_fn():
    """Возвращает текущую функцию логирования."""
    if _log_fn:
        return _log_fn
    # Fallback — print
    return print