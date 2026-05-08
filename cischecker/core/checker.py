"""
Логика проверки кодов маркировки через публичный и True API.
"""
from __future__ import annotations

import json
import sys
import time

from .api import http_post
from .constants import (
    PUBLIC_API, TRUE_API, PG_ALIASES, TIMEOUT, BATCH_SIZE,
)


# ══════════════════════════════════════════════════════════════════════
# Публичный API
# ══════════════════════════════════════════════════════════════════════

def public_check(code: str, debug: bool = False) -> dict | None:
    """Проверка одного кода через публичный API."""
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "okhttp/4.12.0",
    }
    status, body = http_post(PUBLIC_API, json.dumps({"code": code}), headers, debug=debug)
    if status and body:
        try:
            return json.loads(body)
        except json.JSONDecodeError:
            if debug:
                print(f"  [DEBUG] Ответ не JSON: {body[:200]}", file=sys.stderr)
    return None


# ══════════════════════════════════════════════════════════════════════
# True API
# ══════════════════════════════════════════════════════════════════════

def get_pg_from_public(data: dict) -> str:
    """Определяет товарную группу из ответа публичного API."""
    category = data.get("category") or ""
    if not category:
        category = (data.get("codeResolveData") or {}).get("codeCategory", "")
    return category


def resolve_pg_aliases(category: str) -> list[str]:
    """Возвращает варианты pg для True API."""
    category_lower = category.lower()
    if category_lower in PG_ALIASES:
        return PG_ALIASES[category_lower]
    return [category_lower]


def true_check_batch(codes: list[str], pg: str, token: str, debug: bool = False,
                     log_fn=None) -> tuple[int | None, list[dict] | None]:
    """Пакетная проверка через True API. Возвращает (http_status, results)."""
    url = f"{TRUE_API}?pg={pg}"
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
    }

    status, body = http_post(url, json.dumps(codes), headers, debug=debug)

    def _log(msg: str, tag: str = "error"):
        if log_fn:
            log_fn(msg, tag)
        else:
            print(msg, file=sys.stderr)

    if status == 401:
        msg = "❌ Токен недействителен (HTTP 401). Обновите токен."
        if body:
            try:
                err_msg = json.loads(body).get("error_message", "")
                if err_msg:
                    msg = f"❌ Токен недействителен (HTTP 401). Сервер: «{err_msg}»"
            except Exception:
                pass
        _log(msg)
        return (status, None)

    error_map = {
        403: f"❌ Доступ запрещён (HTTP 403). Нет прав для товарной группы «{pg}».",
        429: "⚠ Слишком много запросов (HTTP 429). Превышен лимит — повторите позже.",
        451: "❌ Геоблокировка (HTTP 451). API доступен только с российских IP.",
    }
    if status in error_map:
        _log(error_map[status], "warn" if status == 429 else "error")
        return (status, None)

    if status == 404:
        msg = f"⚠ HTTP 404 для pg={pg} — возможно, неверная товарная группа."
        _log(msg, "warn")
        return (status, None)

    if 500 <= (status or 0) < 600:
        _log(f"❌ Ошибка сервера (HTTP {status}). Попробуйте позже.")
        return (status, None)

    if not status or not body:
        _log("❌ Не удалось подключиться к API. Проверьте интернет и российский IP.")
        return (status, None)

    try:
        data = json.loads(body)
        return (status, data if isinstance(data, list) else [data])
    except json.JSONDecodeError:
        _log(f"⚠ Ответ не JSON: {body[:200]}", "warn")
        return (status, None)


def true_check_with_retry_pg(codes: list[str], category: str, token: str,
                             debug: bool = False, log_fn=None) -> tuple[int | None, list[dict] | None]:
    """Пробует все варианты pg. При 404 — следующий вариант."""
    aliases = resolve_pg_aliases(category)
    for pg in aliases:
        if debug:
            print(f"  [DEBUG] Пробую pg={pg}...", file=sys.stderr)
        status, data = true_check_batch(codes, pg, token, debug=debug, log_fn=log_fn)
        if data is not None:
            return (status, data)
        if status in (401, 403, 429, 451):
            return (status, None)
    return (None, None)


def true_check_auto(codes: list[str], token: str, debug: bool = False,
                    log_fn=None, stop_fn=None, batch_size: int = BATCH_SIZE,
                    progress_fn=None) -> dict[str, dict]:
    """Определяет pg через публичный API, затем True API батчами."""
    from .parser import parse_result  # lazy import to avoid circular
    out: dict[str, dict] = {}

    def _log(msg: str, tag: str = "info"):
        if log_fn:
            log_fn(msg, tag)
        else:
            print(msg, file=sys.stderr)

    def _should_stop():
        return stop_fn() if stop_fn else False

    def _progress(done: int, total: int):
        if progress_fn:
            progress_fn(done, total)

    total_codes = len(codes)
    _log(f"📋 Шаг 1: определение товарных групп для {total_codes} кодов...")
    scanned = 0

    code_to_cat: dict[str, str] = {}
    public_results: dict[str, dict] = {}

    for code in codes:
        if _should_stop():
            _log("⏹ Остановлено на этапе определения товарных групп.", "warn")
            return out
        data = public_check(code, debug=debug)
        scanned += 1
        _progress(scanned, total_codes)
        if data:
            cat = get_pg_from_public(data)
            code_to_cat[code] = cat
            public_results[code] = data
        else:
            out[code] = {"error": "Публичный API недоступен"}
        time.sleep(0.05)  # уменьшено с 0.1

    if not code_to_cat:
        _log("❌ Не удалось определить товарную группу ни для одного кода.", "error")
        return out

    cat_to_codes: dict[str, list[str]] = {}
    for code, cat in code_to_cat.items():
        cat_to_codes.setdefault(cat, []).append(code)

    _log(f"📡 Шаг 2: запрос True API батчами по {batch_size} кодов...")
    checked = 0

    for cat, group in cat_to_codes.items():
        pg_options = resolve_pg_aliases(cat)
        _log(f"  📦 Товарная группа «{cat}»: {len(group)} кодов (pg варианты: {pg_options})")

        for batch_start in range(0, len(group), batch_size):
            if _should_stop():
                _log("⏹ Остановлено на этапе проверки True API.", "warn")
                return out

            batch = group[batch_start:batch_start + batch_size]
            batch_num = batch_start // batch_size + 1
            total_batches = (len(group) + batch_size - 1) // batch_size
            _log(f"    Батч {batch_num}/{total_batches} ({len(batch)} кодов)...")

            status, data = true_check_with_retry_pg(batch, cat, token, debug=debug, log_fn=log_fn)

            if data:
                found = set()
                for item in data:
                    cis_info = item.get("cisInfo", item)
                    c = cis_info.get("requestedCis", cis_info.get("cis", ""))
                    if c:
                        found.add(c)
                    out[c] = item
                for code in batch:
                    if code not in found:
                        out[code] = out.get(code, {})
                        if "error" not in out[code]:
                            out[code]["_public_data"] = public_results.get(code)
                checked += len(batch)
                _log(f"    ✓ Батч {batch_num}/{total_batches}: получено {len(data)} результатов", "success")
            else:
                error_msg = explain_http_status(status)
                for code in batch:
                    out[code] = out.get(code, {"error": f"True API: {error_msg}"})
                checked += len(batch)
                _log(f"    ✗ Батч {batch_num}/{total_batches}: {error_msg}", "error")

            _progress(scanned + checked, total_codes)
            time.sleep(0.05)  # уменьшено с 0.15

    return out


def explain_http_status(status: int | None) -> str:
    """Расшифровка HTTP-статуса."""
    if status is None:
        return "нет соединения с сервером"
    mapping = {
        400: "некорректный запрос",
        401: "токен недействителен или просрочен",
        403: "доступ запрещён",
        404: "код не найден или неверная товарная группа",
        429: "превышен лимит запросов",
        451: "геоблокировка — нужен российский IP",
    }
    if status in mapping:
        return f"HTTP {status}: {mapping[status]}"
    if 500 <= status < 600:
        return f"HTTP {status}: ошибка сервера"
    return f"HTTP {status}"
