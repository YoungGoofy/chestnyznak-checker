"""
HTTP-клиент с поддержкой keep-alive для повторного использования TCP-соединений.
"""
from __future__ import annotations

import json
import http.client
import ssl
import sys
import time
from urllib.parse import urlparse

from .constants import TIMEOUT, RETRY, RETRY_DELAY


class ApiSession:
    """HTTP-сессия с keep-alive для одного хоста.

    Переиспользует TCP+TLS соединение между запросами,
    экономя ~200-400ms на каждый запрос (TLS handshake).
    """

    def __init__(self, host: str, timeout: int = TIMEOUT):
        self._host = host
        self._timeout = timeout
        self._conn: http.client.HTTPSConnection | None = None
        self._context = ssl.create_default_context()

    def _get_conn(self) -> http.client.HTTPSConnection:
        """Возвращает живое соединение (или создаёт новое)."""
        if self._conn is not None:
            try:
                # Проверяем, что соединение живо
                self._conn.request("HEAD", "/", headers={"Host": self._host})
                self._conn.getresponse()
                return self._conn
            except Exception:
                self._close()

        self._conn = http.client.HTTPSConnection(
            self._host, timeout=self._timeout, context=self._context,
        )
        return self._conn

    def _close(self) -> None:
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None

    def post(self, path: str, payload: str, headers: dict,
             debug: bool = False) -> tuple[int | None, str | None]:
        """POST с ретраями и keep-alive."""
        for attempt in range(1, RETRY + 1):
            try:
                conn = self._get_conn()
                conn.request("POST", path, body=payload.encode("utf-8"), headers=headers)
                resp = conn.getresponse()
                body = resp.read().decode("utf-8")

                if debug:
                    print(f"\n  [DEBUG] POST https://{self._host}{path}", file=sys.stderr)
                    print(f"  [DEBUG] HTTP {resp.status}, body: {body[:2000]}", file=sys.stderr)

                return (resp.status, body)

            except http.client.HTTPException as e:
                self._close()
                if debug:
                    print(f"  [DEBUG] HTTP exception (attempt {attempt}): {e}", file=sys.stderr)
                if attempt == RETRY:
                    return (None, str(e))
                time.sleep(RETRY_DELAY * attempt)

            except Exception as e:
                self._close()
                if debug:
                    print(f"  [DEBUG] Exception (attempt {attempt}): {e}", file=sys.stderr)
                if attempt == RETRY:
                    return (None, str(e))
                time.sleep(RETRY_DELAY * attempt)

        return (None, None)

    def get_json(self, path: str, headers: dict | None = None) -> dict | None:
        """GET-запрос, возвращает JSON."""
        hdrs = {"Accept": "application/json", "Host": self._host}
        if headers:
            hdrs.update(headers)
        try:
            conn = self._get_conn()
            conn.request("GET", path, headers=hdrs)
            resp = conn.getresponse()
            body = resp.read().decode("utf-8")
            if resp.status == 200:
                return json.loads(body)
        except Exception:
            self._close()
        return None

    def __del__(self):
        self._close()


# ── Глобальные сессии (переиспользуются между запросами) ────────────────

_sessions: dict[str, ApiSession] = {}


def _get_session(host: str) -> ApiSession:
    """Возвращает или создаёт сессию для хоста."""
    if host not in _sessions:
        _sessions[host] = ApiSession(host)
    return _sessions[host]


def http_post(url: str, payload_str: str, headers: dict,
              debug: bool = False) -> tuple[int | None, str | None]:
    """POST-запрос с keep-alive. Drop-in замена для оригинальной http_post()."""
    parsed = urlparse(url)
    session = _get_session(parsed.hostname or "")
    path = parsed.path
    if parsed.query:
        path += "?" + parsed.query

    # Добавляем Host заголовок
    hdrs = dict(headers)
    hdrs.setdefault("Host", parsed.hostname or "")

    return session.post(path, payload_str, hdrs, debug=debug)
