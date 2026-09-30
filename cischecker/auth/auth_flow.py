"""
Авторизация через УКЭП: единый токен аутентификации в формате UUID.

Основной поток (раздел 1.5.2 True API, docs.crpt.ru/gismt/True_API/):
  подпись ИНН участника оборота (attached CMS, base64)
  → POST /auth/simpleSignIn {"data": ..., "unitedToken": true}
  → {"uuidToken": "...", "expireDate": "yyyy-MM-ddTHH:mm:ss.SSSZ"}

Фолбэк: классическая пара из GET /auth/key (uuid + подпись challenge)
с тем же флагом unitedToken.

Прежний JWT-поток удалён: JWT — временно поддерживаемый ЧЗ формат
(до 31.12.2026), основной формат — UUID.
"""
from __future__ import annotations

import base64
import json
import os
from datetime import datetime, timezone
from urllib.request import urlopen, Request
from urllib.error import HTTPError

from ..core.constants import AUTH_KEY_URL, AUTH_SIGN_URL, TIMEOUT
from .signer import sign_data

# Логирование
_log_fn = None

_SIGN_ERROR = (
    "Не удалось подписать данные.\n\n"
    "Проверьте:\n"
    "• КриптоПро CSP установлен (версия 5.x рекомендуется)\n"
    "• USB-токен (RuToken/eToken) подключён\n"
    "• Сертификат установлен в хранилище «Личные» (My)\n"
    "• pywin32 установлен (pip install pywin32)\n\n"
    "Для диагностики откройте КриптоПро CSP → Сервис → Просмотреть сертификаты."
)


def set_log_fn(fn) -> None:
    global _log_fn
    _log_fn = fn


def _log(msg: str, tag: str = "info") -> None:
    if _log_fn:
        _log_fn(msg, tag)
    else:
        print(msg)


def _iso_to_ts(iso_str: str | None) -> float | None:
    """'2026-10-10T00:00:00.123Z' → unix timestamp."""
    if not iso_str:
        return None
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except ValueError:
        return None


def _fetch_challenge() -> tuple[str, str] | None:
    """GET /auth/key → (uuid, data)."""
    _log(f"📡 Запрашиваю пару авторизации: {AUTH_KEY_URL}")
    try:
        req = Request(AUTH_KEY_URL, headers={"Accept": "application/json"})
        with urlopen(req, timeout=TIMEOUT) as resp:
            challenge = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        _log(f"   ❌ Не удалось получить пару авторизации: {e}", "error")
        return None

    uuid = challenge.get("uuid", "")
    data_str = challenge.get("data", "")
    if not uuid or not data_str:
        _log(f"   ❌ Некорректный ответ: {json.dumps(challenge, ensure_ascii=False)[:200]}", "error")
        return None
    _log(f"   ✅ Пара получена (uuid: {uuid[:12]}...)")
    return uuid, data_str


def _post_sign_in(body: dict) -> tuple[int | None, dict | str]:
    """POST /auth/simpleSignIn → (http_status, parsed_json | error)."""
    _log(f"📡 Отправляю подпись: {AUTH_SIGN_URL}")
    req = Request(AUTH_SIGN_URL, data=json.dumps(body).encode("utf-8"), headers={
        "Content-Type": "application/json",
        "Accept": "application/json",
    })
    try:
        with urlopen(req, timeout=TIMEOUT) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except HTTPError as e:
        try:
            err_body = json.loads(e.read().decode("utf-8"))
        except Exception:
            err_body = str(e)
        return e.code, err_body
    except Exception as e:
        return None, str(e)


def _extract_token(result) -> tuple[str, float | None]:
    """Ответ → (token, expires_ts). Понимает uuidToken и legacy token."""
    if isinstance(result, dict):
        token = result.get("uuidToken") or result.get("token") or ""
        return token, _iso_to_ts(result.get("expireDate"))
    return "", None


def auth_uuid_token(thumbprint: str = "", inn: str = "") -> tuple[bool, str, float | None]:
    """
    Получает единый токен (UUID) через УКЭП.

    Возвращает (success: bool, token_or_error: str, expires_ts: float | None).
    """
    inn = (inn or os.environ.get("CHESTNYZNAK_INN", "")).strip()
    if not inn:
        return False, (
            "Не удалось определить ИНН участника оборота.\n\n"
            "Для единого токена подписывается ИНН организации:\n"
            "• ИНН берётся из сертификата УКЭП (OID 1.2.643.3.131.1.1)\n"
            "• либо задайте CHESTNYZNAK_INN в .env"
        ), None

    # ── Попытка 1: подпись ИНН, без предварительного /auth/key ──
    _log(f"🔐 Подписываю ИНН ({inn}) сертификатом УКЭП...")
    if thumbprint:
        _log(f"   Thumbprint: {thumbprint[:16]}...")

    signature = sign_data(inn.encode("ascii"), thumbprint)
    if signature is None:
        return False, _SIGN_ERROR, None
    _log("   ✅ Подпись создана")

    body = {
        "data": base64.b64encode(signature).decode("ascii"),
        "unitedToken": True,
    }
    status, result = _post_sign_in(body)
    token, expires_ts = _extract_token(result)

    if not token:
        # ── Попытка 2: пара uuid + challenge из /auth/key ──
        _log("🔁 Основной поток не принят, пробую через /auth/key (uuid + challenge)...", "warn")
        challenge = _fetch_challenge()
        if challenge is None:
            return False, _parse_auth_error(status, result), None

        uuid, data_str = challenge
        signature = sign_data(data_str.encode("utf-8"), thumbprint)
        if signature is None:
            return False, _SIGN_ERROR, None

        body = {
            "uuid": uuid,
            "data": base64.b64encode(signature).decode("ascii"),
            "unitedToken": True,
        }
        status, result = _post_sign_in(body)
        token, expires_ts = _extract_token(result)
        if not token:
            error_msg = _parse_auth_error(status, result)
            _log(f"   ❌ {error_msg}", "error")
            return False, error_msg, None

    _log("   ✅ Единый токен (UUID) получен!", "success")
    return True, token, expires_ts


def _parse_auth_error(status: int | None, response) -> str:
    """Расшифровка ошибок авторизации."""
    if isinstance(response, dict):
        msg = response.get("error_message", "") or response.get("message", "")
        if msg:
            return f"Ошибка авторизации (HTTP {status}): {msg}"
    if isinstance(response, str):
        return f"Ошибка авторизации (HTTP {status}): {response}"
    if status == 400:
        return "Ошибка запроса (HTTP 400). Проверьте формат подписи."
    if status == 403:
        return "Доступ запрещён (HTTP 403). Пользователь не найден или неактивен в ЧЗ."
    return f"Ошибка авторизации (HTTP {status})."
