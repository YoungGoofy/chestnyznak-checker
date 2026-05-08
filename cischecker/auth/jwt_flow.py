"""
JWT-авторизация через УКЭП (2-шаговый flow).

Шаг 1: GET /auth/key → uuid + data (challenge)
Шаг 2: sign(data) → POST /auth/simpleSignIn → JWT-токен
"""
from __future__ import annotations

import base64
import json
from urllib.request import urlopen, Request
from urllib.error import HTTPError

from ..core.constants import AUTH_KEY_URL, AUTH_SIGN_URL, TIMEOUT
from .signer import sign_data

# Логирование
_log_fn = None


def set_log_fn(fn) -> None:
    global _log_fn
    _log_fn = fn


def _log(msg: str, tag: str = "info") -> None:
    if _log_fn:
        _log_fn(msg, tag)
    else:
        print(msg)


def auth_jwt(thumbprint: str = "") -> tuple[bool, str]:
    """
    Получает JWT-токен через двухшаговый flow.

    Возвращает (success: bool, token_or_error: str).
    """
    # ── Шаг 1: Получаем challenge ──
    _log("📡 Шаг 1: получаю challenge с сервера ЧЗ...")
    _log(f"   URL: {AUTH_KEY_URL}")

    try:
        req = Request(AUTH_KEY_URL, headers={"Accept": "application/json"})
        with urlopen(req, timeout=TIMEOUT) as resp:
            challenge = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return False, f"Не удалось получить challenge: {e}"

    uuid = challenge.get("uuid", "")
    data_str = challenge.get("data", "")

    if not uuid or not data_str:
        return False, f"Сервер вернул некорректный challenge: {json.dumps(challenge, ensure_ascii=False)[:200]}"

    _log(f"   ✅ Challenge получен (uuid: {uuid[:12]}...)")

    # ── Шаг 2: Подписываем challenge ──
    _log("🔐 Шаг 2: подписываю challenge сертификатом УКЭП...")
    if thumbprint:
        _log(f"   Thumbprint: {thumbprint[:16]}...")

    signature = sign_data(data_str.encode("utf-8"), thumbprint)
    if signature is None:
        return False, (
            "Не удалось подписать данные.\n\n"
            "Проверьте:\n"
            "• КриптоПро CSP установлен (версия 5.x рекомендуется)\n"
            "• USB-токен (RuToken/eToken) подключён\n"
            "• Сертификат установлен в хранилище «Личные» (My)\n"
            "• pywin32 установлен (pip install pywin32)\n\n"
            "Для диагностики откройте КриптоПро CSP → Сервис → Просмотреть сертификаты."
        )

    sig_b64 = base64.b64encode(signature).decode("ascii")
    _log(f"   ✅ Подпись создана ({len(signature)} байт)")

    # ── Шаг 3: Отправляем подпись ──
    _log("📡 Шаг 3: отправляю подпись на сервер...")
    _log(f"   URL: {AUTH_SIGN_URL}")

    body = json.dumps({"uuid": uuid, "data": sig_b64}).encode("utf-8")
    req = Request(AUTH_SIGN_URL, data=body, headers={
        "Content-Type": "application/json",
        "Accept": "application/json",
    })

    try:
        with urlopen(req, timeout=TIMEOUT) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            status = resp.status
    except HTTPError as e:
        try:
            err_body = json.loads(e.read().decode("utf-8"))
        except Exception:
            err_body = str(e)
        error_msg = _parse_auth_error(e.code, err_body)
        _log(f"   ❌ {error_msg}", "error")
        return False, error_msg
    except Exception as e:
        return False, f"Ошибка при отправке подписи: {e}"

    if status == 200 and isinstance(result, dict):
        token = result.get("token", "")
        if token:
            _log("   ✅ JWT-токен получен!", "success")
            return True, token

    error_msg = _parse_auth_error(status, result)
    _log(f"   ❌ {error_msg}", "error")
    return False, error_msg


def _parse_auth_error(status: int, response) -> str:
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
