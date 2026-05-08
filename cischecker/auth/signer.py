"""
Подпись данных УКЭП (Windows only).

Порядок методов:
1. CAdESCOM (КриптоПро CSP 5.x) — CAdESCOM.CPSigner + CAdESCOM.CadesSignedData
2. Legacy COM (CSP 4.x) — CPSigner + CPSignedData
"""
from __future__ import annotations

import base64
import platform

from .certificates import (
    find_cert_in_store,
    CAPICOM_CURRENT_USER_STORE,
    CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED,
    CAPICOM_STORE_OPEN_READ_ONLY,
)

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


def sign_data(data: bytes, thumbprint: str = "") -> bytes | None:
    """
    Подписывает данные УКЭП (attached CMS, base64-encoded DER).

    Возвращает CMS signature bytes или None при ошибке.
    Только Windows.
    """
    if platform.system() != "Windows":
        _log("❌ Подпись УКЭП доступна только на Windows.", "error")
        return None

    try:
        import win32com.client  # noqa: F401
    except ImportError:
        _log("❌ pywin32 не установлен. Установите: pip install pywin32", "error")
        return None

    # 1. CAdESCOM (CSP 5.x)
    result = _sign_cadescom(data, thumbprint)
    if result is not None:
        return result

    # 2. Legacy COM (CSP 4.x)
    result = _sign_legacy_com(data, thumbprint)
    if result is not None:
        return result

    _log("❌ Ни один метод подписи не сработал (CAdESCOM, CPCSPStore)", "error")
    return None


def _sign_cadescom(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через CAdESCOM (КриптоПро CSP 5.x).

    Хранилище: CAdESCOM.Store или CAPICOM.Store.
    Signer: CAdESCOM.CPSigner
    Data: CAdESCOM.CadesSignedData
    """
    try:
        import win32com.client
    except ImportError:
        return None

    # Пробуем открыть хранилище: сначала CAdESCOM.Store, потом CAPICOM.Store
    store = None
    for store_progid in ("CAdESCOM.Store", "CAPICOM.Store"):
        try:
            store = win32com.client.Dispatch(store_progid)
            store.Open(
                CAPICOM_CURRENT_USER_STORE,
                "My",
                CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED,
            )
            _log(f"  🔐 Хранилище {store_progid} открыто для подписи")
            break
        except Exception:
            store = None
            continue

    if store is None:
        _log("  ⚠ Не удалось открыть хранилище сертификатов для подписи")
        return None

    try:
        # Находим сертификат
        cert = None
        if thumbprint:
            cert = find_cert_in_store(store, thumbprint)
            if cert:
                _log("  ✅ Сертификат найден по thumbprint")
            else:
                _log("  ⚠ Сертификат с указанным thumbprint не найден, пробуем первый...")

        if cert is None:
            count = store.Certificates.Count
            if count > 0:
                cert = store.Certificates.Item(1)
                _log(f"  ℹ Используется первый сертификат из {count}")
            else:
                _log("  ⚠ Хранилище пусто — нет сертификатов для подписи")
                store.Close()
                return None

        store.Close()

        # Создаём Signer
        signer = win32com.client.Dispatch("CAdESCOM.CPSigner")
        signer.Certificate = cert

        # Создаём SignedData
        signed_data = win32com.client.Dispatch("CAdESCOM.CadesSignedData")
        # Content — строка с данными для подписи (не base64!)
        signed_data.Content = base64.b64encode(data).decode("ascii")

        # Подписываем: Sign(Signer, Detached, EncodingType)
        # Detached=False → attached CMS
        # EncodingType: 0=Base64
        signature_b64 = signed_data.Sign(signer, False, 0)

        _log("  ✅ Данные подписаны (CAdESCOM)", "success")
        return base64.b64decode(signature_b64)

    except Exception as e:
        _log(f"  ⚠ CAdESCOM подпись: {e}")
        try:
            store.Close()
        except Exception:
            pass
        return None


def _sign_legacy_com(data: bytes, thumbprint: str = "") -> bytes | None:
    """Подпись через legacy COM (КриптоПро CSP 4.x)."""
    try:
        import win32com.client
    except ImportError:
        return None

    try:
        store = win32com.client.Dispatch("CPCSPStore.Store")
        store.Open()

        cert = None
        if thumbprint:
            cert = find_cert_in_store(store, thumbprint)
        if cert is None and store.Certificates.Count > 0:
            cert = store.Certificates.Item(1)

        if not cert:
            store.Close()
            _log("  ⚠ Сертификат не найден (CPCSPStore)")
            return None

        signer = win32com.client.Dispatch("CPSigner.Signer")
        signer.Certificate = cert

        signed_data = win32com.client.Dispatch("CPSignedData.SignedData")
        signed_data.Content = base64.b64encode(data).decode("ascii")

        signature_b64 = signed_data.SignCPS(signer, True, 0)

        store.Close()
        _log("  ✅ Данные подписаны (CPCSPStore)", "success")
        return base64.b64decode(signature_b64)

    except Exception as e:
        _log(f"  ⚠ CPCSPStore подпись: {e}")
        return None
