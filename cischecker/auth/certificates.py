"""
Обнаружение сертификатов УКЭП (Windows only).

Порядок поиска:
1. CAdESCOM.Store (КриптоПро CSP 5.x — основной метод)
2. CAPICOM.Store (Windows Certificate Store, перехватывается КриптоПро)
3. CPCSPStore.Store (legacy CSP 4.x)

ВАЖНО: КриптоПро CSP 5.x предоставляет CAdESCOM.Store.
       CAPICOM.Store может быть недоступен, если не установлен CAPICOM SDK.

ВАЖНО: Все функции, использующие COM-объекты, должны вызывать
       pythoncom.CoInitialize() / CoUninitialize() для корректной работы
       в фоновых потоках (threading.Thread). Без этого вызова COM-объекты
       недоступны и возвращают ошибку -2147221008 (CoInitialize not called).
"""
from __future__ import annotations

import contextlib
import platform

# ── Константы хранилищ ─────────────────────────────────────────────────
CAPICOM_CURRENT_USER_STORE = 2
CAPICOM_STORE_OPEN_READ_ONLY = 0
CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = 2

# OID КриптоПро для ИНН
OID_INN = "1.2.643.3.131.1.1"

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


@contextlib.contextmanager
def _com_initialized():
    """Контекстный менеджер для инициализации COM в текущем потоке.

    Необходим для корректной работы COM-объектов в фоновых потоках
    (threading.Thread). Без CoInitialize() COM возвращает ошибку
    -2147221008 (CoInitialize was not called).
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            yield
        finally:
            pythoncom.CoUninitialize()
    except ImportError:
        # pythoncom недоступен (не Windows или нет pywin32)
        yield


def list_certificates() -> list[dict]:
    """
    Возвращает список сертификатов УКЭП.

    Каждый сертификат: {
        "thumbprint": "XX XX ...",
        "subject": "...",
        "issuer": "...",
        "not_before": "...",
        "not_after": "...",
        "inn": "1234567890",
    }
    """
    if platform.system() != "Windows":
        _log("⚠ УКЭП доступен только на Windows.", "warn")
        return []

    try:
        import win32com.client  # noqa: F401
    except ImportError:
        _log(
            "❌ Библиотека pywin32 не установлена.\n"
            "   Установите: pip install pywin32\n"
            "   Затем перезапустите приложение.",
            "error",
        )
        return []

    with _com_initialized():

        # 1. CAdESCOM.Store — КриптоПро CSP 5.x (основной)
        certs = _list_certs_cadescom_store()
        if certs:
            _log(f"📋 Найдено сертификатов (CAdESCOM.Store): {len(certs)}", "success")
            return certs

        # 2. CAPICOM.Store — Windows Certificate Store
        certs = _list_certs_capicom_store()
        if certs:
            _log(f"📋 Найдено сертификатов (CAPICOM.Store): {len(certs)}", "success")
            return certs

        # 3. CPCSPStore — legacy CSP 4.x
        certs = _list_certs_legacy_store()
        if certs:
            _log(f"📋 Найдено сертификатов (CPCSPStore): {len(certs)}", "success")
            return certs

        _log(
            "⚠ Сертификаты не найдены.\n"
            "   Проверьте:\n"
            "   • КриптоПро CSP установлен и лицензия активна\n"
            "   • USB-токен (RuToken/eToken) подключён\n"
            "   • Сертификат установлен в хранилище «Личные»\n"
            "   • В КриптоПро CSP → Сервис → Просмотреть сертификаты — сертификат виден",
            "warn",
        )
        return []


def _list_certs_cadescom_store() -> list[dict]:
    """Сертификаты через CAdESCOM.Store (КриптоПро CSP 5.x).

    Это основной метод для CSP 5.x. ProgID 'CAdESCOM.Store'
    доступен после установки КриптоПро CSP 5.x.
    """
    try:
        import win32com.client
        store = win32com.client.Dispatch("CAdESCOM.Store")
    except Exception as e:
        _log(f"  ℹ CAdESCOM.Store недоступен: {e}")
        return []

    certs = []
    try:
        store.Open(
            CAPICOM_CURRENT_USER_STORE,
            "My",
            CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED,
        )
        _log(f"  📂 CAdESCOM.Store открыт (хранилище «Личные»)")

        # Count — это СВОЙСТВО, не метод (без скобок!)
        count = store.Certificates.Count
        _log(f"  📂 Сертификатов в хранилище: {count}")

        for i in range(1, count + 1):
            try:
                cert = store.Certificates.Item(i)
                info = _parse_cert(cert)
                if info:
                    certs.append(info)
            except Exception as e:
                _log(f"  ⚠ Ошибка чтения сертификата #{i}: {e}")

        store.Close()
    except Exception as e:
        _log(f"  ⚠ CAdESCOM.Store ошибка: {e}")
        try:
            store.Close()
        except Exception:
            pass

    return certs


def _list_certs_capicom_store() -> list[dict]:
    """Сертификаты через CAPICOM.Store (Windows standard).

    CAPICOM.Store — стандартный Windows COM ProgID.
    КриптоПро CSP перехватывает его и добавляет ГОСТ-сертификаты.
    Может быть недоступен на свежих Windows без CAPICOM SDK.
    """
    try:
        import win32com.client
        store = win32com.client.Dispatch("CAPICOM.Store")
    except Exception as e:
        _log(f"  ℹ CAPICOM.Store недоступен: {e}")
        return []

    certs = []
    try:
        store.Open(
            CAPICOM_CURRENT_USER_STORE,
            "My",
            CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED,
        )

        count = store.Certificates.Count
        _log(f"  📂 CAPICOM.Store: {count} сертификатов")

        for i in range(1, count + 1):
            try:
                cert = store.Certificates.Item(i)
                info = _parse_cert(cert)
                if info:
                    certs.append(info)
            except Exception as e:
                _log(f"  ⚠ Ошибка сертификата #{i}: {e}")

        store.Close()
    except Exception as e:
        _log(f"  ⚠ CAPICOM.Store ошибка: {e}")
        try:
            store.Close()
        except Exception:
            pass

    return certs


def _list_certs_legacy_store() -> list[dict]:
    """Сертификаты через CPCSPStore (КриптоПро CSP 4.x legacy)."""
    try:
        import win32com.client
        store = win32com.client.Dispatch("CPCSPStore.Store")
    except Exception:
        return []

    certs = []
    try:
        store.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_READ_ONLY)

        for cert in store.Certificates:
            info = _parse_cert(cert)
            if info:
                certs.append(info)

        store.Close()
    except Exception as e:
        _log(f"  ⚠ CPCSPStore: {e}")
        try:
            store.Close()
        except Exception:
            pass

    return certs


def _parse_cert(cert) -> dict | None:
    """Парсит COM-объект сертификата в словарь."""
    try:
        # SubjectName (КриптоПро) или Subject (CAPICOM standard)
        subject = ""
        for attr in ("SubjectName", "Subject"):
            try:
                subject = getattr(cert, attr, "") or ""
                if subject:
                    break
            except Exception:
                continue

        # IssuerName или Issuer
        issuer = ""
        for attr in ("IssuerName", "Issuer"):
            try:
                issuer = getattr(cert, attr, "") or ""
                if issuer:
                    break
            except Exception:
                continue

        thumbprint = ""
        try:
            thumbprint = cert.Thumbprint or ""
        except Exception:
            pass

        not_before = ""
        try:
            not_before = str(cert.ValidFromDate or "")
        except Exception:
            pass

        not_after = ""
        try:
            not_after = str(cert.ValidToDate or "")
        except Exception:
            pass

        # Извлекаем ИНН
        inn = _extract_inn_from_subject(subject)

        if not thumbprint:
            return None

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


def _extract_inn_from_subject(subject: str) -> str:
    """Извлекает ИНН из Subject DN сертификата."""
    if not subject:
        return ""

    for sep in [", ", "; ", "\n"]:
        parts = subject.split(sep)
        for part in parts:
            part = part.strip()
            # INN=... / ИНН=...
            if part.upper().startswith("INN=") or part.upper().startswith("ИНН="):
                val = part.split("=", 1)[1].strip()
                if val.isdigit() and len(val) in (10, 12):
                    return val
            # OID: 1.2.643.3.131.1.1=#hex
            if part.startswith(OID_INN + "="):
                val = part.split("=", 1)[1].strip()
                if val.startswith("#"):
                    try:
                        hex_str = val[1:]
                        if len(hex_str) >= 8:
                            inn_hex = hex_str[4:]  # пропускаем ASN.1 tag+length
                            inn_val = bytes.fromhex(inn_hex).decode("ascii", errors="ignore").strip()
                            if inn_val.isdigit() and len(inn_val) in (10, 12):
                                return inn_val
                    except Exception:
                        pass
                elif val.isdigit() and len(val) in (10, 12):
                    return val
    return ""


def find_cert_in_store(store, thumbprint: str):
    """Находит сертификат по thumbprint в открытом COM Store."""
    tp_clean = thumbprint.lower().replace(" ", "")
    for cert in store.Certificates:
        try:
            cert_tp = (cert.Thumbprint or "").lower().replace(" ", "")
            if cert_tp and cert_tp == tp_clean:
                return cert
        except Exception:
            continue
    return None


def diagnose_com() -> str:
    """Диагностика: какие COM-объекты КриптоПро доступны.

    Вызывается для отладки, если сертификаты не найдены.
    Использует CoInitialize() для работы в фоновых потоках.
    """
    if platform.system() != "Windows":
        return "Диагностика доступна только на Windows."

    try:
        import win32com.client
    except ImportError:
        return "pywin32 НЕ установлен. Установите: pip install pywin32"

    with _com_initialized():
        results = []

        for prog_id in ["CAdESCOM.Store", "CAPICOM.Store", "CPCSPStore.Store",
                         "CAdESCOM.CPSigner", "CAdESCOM.CadesSignedData",
                         "CAdESCOM.About"]:
            try:
                obj = win32com.client.Dispatch(prog_id)
                results.append(f"  ✅ {prog_id} — доступен")
                # Для CAdESCOM.About можно получить версию
                if prog_id == "CAdESCOM.About":
                    try:
                        ver = obj.Version
                        results.append(f"     Версия КриптоПро: {ver}")
                    except Exception:
                        pass
            except Exception as e:
                results.append(f"  ❌ {prog_id} — НЕДОСТУПЕН ({e})")

        return "Диагностика COM-объектов КриптоПро:\n" + "\n".join(results)
