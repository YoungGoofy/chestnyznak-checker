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

Подавление UI-окон КриптоПро (silent mode):
  Обращение к cert.PrivateKey.UniqueContainerName через COM вызывает CSP
  Provider, который показывает системный диалог «выберите считыватель»
  для сертификатов из кэша с отключённым токеном.

  Стратегия: сначала читаем CERT_KEY_PROV_INFO_PROP_ID через CryptoAPI
  (без COM, без UI), затем проверяем физическую доступность ключа через
  CryptAcquireContextW с флагом CRYPT_SILENT. Если ключ недоступен —
  молча пропускаем сертификат. Если доступен — безопасно читаем
  cert.PrivateKey.UniqueContainerName (UI не появится, т.к. токен вставлен).

Дедупликация:
  Сертификаты собираются из ВСЕХ доступных Store (CAdESCOM + CAPICOM +
  CPCSPStore), затем дедуплицируются по thumbprint, чтобы исключить
  дубли из разных хранилищ.
"""
from __future__ import annotations

import contextlib
import platform
from datetime import datetime

# ── Константы хранилищ ─────────────────────────────────────────────────
CAPICOM_CURRENT_USER_STORE = 2
CAPICOM_STORE_OPEN_READ_ONLY = 0
CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = 2

# OID КриптоПро для ИНН
OID_INN = "1.2.643.3.131.1.1"

# ── Константы CryptoAPI ────────────────────────────────────────────────
CERT_KEY_PROV_INFO_PROP_ID = 2
CRYPT_SILENT = 0x40
CRYPT_MACHINE_KEYSET = 0x20

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


def _is_expired(not_after: str) -> bool:
    """Проверяет, просрочен ли сертификат по дате ValidToDate."""
    if not not_after:
        return False  # Не можем определить — показываем
    try:
        # КриптоПро возвращает даты в разных форматах
        for fmt in (
            "%m/%d/%Y %I:%M:%S %p",    # 01/15/2027 12:00:00 PM (EN)
            "%d.%m.%Y %H:%M:%S",       # 15.01.2027 12:00:00 (RU)
            "%Y-%m-%dT%H:%M:%S",       # ISO
            "%Y-%m-%d %H:%M:%S",       # ISO space
            "%m/%d/%Y %H:%M:%S",       # US 24h
            "%d.%m.%Y",               # RU short
        ):
            try:
                expiry = datetime.strptime(not_after.strip(), fmt)
                return expiry < datetime.now()
            except ValueError:
                continue
        # Fallback: попробовать через dateutil-подобный парсинг
        # Если не удалось распознать — не фильтруем
        return False
    except Exception:
        return False


def _is_on_removable_media(container_name: str) -> bool:
    r"""Определяет, находится ли закрытый ключ на съёмном носителе.

    КриптоПро UniqueContainerName содержит путь вида:
      \\.\REGISTRY\...  — реестр (ЛОКАЛЬНОЕ хранилище)
      \\.\HDIMAGE\...   — виртуальный жёсткий диск КриптоПро (ЛОКАЛЬНОЕ)
      \\.\FLASH\...     — USB-флешка (СЪЁМНЫЙ)
      \\.\FAT12\...     — FAT12, eToken/RuToken (СЪЁМНЫЙ)
      \\.\FAT16\...     — FAT16 (СЪЁМНЫЙ)
      \\.\aktiv co. ruToken 0\... — RuToken (СЪЁМНЫЙ)
      \\.\JaCarta\...   — JaCarta (СЪЁМНЫЙ)
      \\.\eToken\...    — eToken (СЪЁМНЫЙ)
      \\.\<reader>\...  — любой другой считыватель (СЪЁМНЫЙ)

    Правило: REGISTRY и HDIMAGE — локальные. Всё остальное с \\.\ —
    считыватель/носитель, т.е. съёмный.
    """
    if not container_name:
        return False
    cn_upper = container_name.upper()
    # Локальные хранилища КриптоПро — точно НЕ съёмный носитель
    if cn_upper.startswith("\\\\.\\REGISTRY") or cn_upper.startswith("\\\\.\\HDIMAGE"):
        return False
    # Любой \\.\<reader> — считыватель, съёмный носитель
    if cn_upper.startswith("\\\\.\\"):
        return True
    # Нестандартный формат (логическое имя контейнера без пути) —
    # не можем определить. Считаем съёмным, т.к. ключ уже проверен
    # на доступность через _test_key_accessible_silent.
    return True


# ═══════════════════════════════════════════════════════════════════════
# Silent-проверки ключей (без UI-диалогов КриптоПро)
# ═══════════════════════════════════════════════════════════════════════

def _read_cert_prov_info(cert) -> dict | None:
    """Читает CERT_KEY_PROV_INFO_PROP_ID из сертификата через CryptoAPI.

    НЕ обращается к cert.PrivateKey — не вызывает UI-диалоги.

    Возвращает dict с ключами:
      ContainerName — логическое имя контейнера (напр. "lea-0c5e0...")
      ProvName       — имя CSP (напр. "Crypto-Pro GOST R 34.10-2012 KC1 Strong CSP")
      ProvType       — тип CSP (целое)
      Flags          — флаги (CRYPT_MACHINE_KEYSET и пр.)
    Или None, если свойство недоступно.
    """
    cert_handle = None
    try:
        cert_handle = cert.Handle
    except Exception:
        return None

    if not cert_handle:
        return None

    # Способ 1: win32crypt (если cert.Handle совместим с PyCERT_CONTEXT)
    try:
        import win32crypt
        import win32cryptcon
        prov_info = win32crypt.CertGetCertificateContextProperty(
            cert_handle, win32cryptcon.CERT_KEY_PROV_INFO_PROP_ID
        )
        if prov_info:
            return {
                "ContainerName": prov_info.get("ContainerName", ""),
                "ProvName": prov_info.get("ProvName", ""),
                "ProvType": prov_info.get("ProvType", 0),
                "Flags": prov_info.get("Flags", 0),
            }
    except Exception:
        pass  # cert.Handle несовместим с win32crypt — пробуем ctypes

    # Способ 2: ctypes crypt32.CertGetCertificateContextProperty
    try:
        handle_int = int(cert_handle)
        if handle_int:
            return _read_prov_info_ctypes(handle_int)
    except (TypeError, ValueError):
        pass

    return None


def _read_prov_info_ctypes(cert_handle_int: int) -> dict | None:
    """Читает CRYPT_KEY_PROV_INFO через ctypes (crypt32.dll).

    cert_handle_int: целочисленное значение PCCERT_CONTEXT.
    """
    import ctypes
    from ctypes import wintypes

    crypt32 = ctypes.WinDLL("crypt32")

    cb_data = wintypes.DWORD(0)
    if not crypt32.CertGetCertificateContextProperty(
        cert_handle_int, CERT_KEY_PROV_INFO_PROP_ID,
        None, ctypes.byref(cb_data),
    ) or cb_data.value == 0:
        return None

    buf = ctypes.create_string_buffer(cb_data.value)
    if not crypt32.CertGetCertificateContextProperty(
        cert_handle_int, CERT_KEY_PROV_INFO_PROP_ID,
        buf, ctypes.byref(cb_data),
    ):
        return None

    # Структура CRYPT_KEY_PROV_INFO:
    #   LPCWSTR pwszContainerName;   // ptr [0 .. ptr_size)
    #   LPCWSTR pwszProvName;        // ptr [ptr_size .. 2*ptr_size)
    #   DWORD   dwProvType;          // [2*ptr_size .. 2*ptr_size+4)
    #   DWORD   dwFlags;             // [2*ptr_size+4 .. 2*ptr_size+8)
    #   DWORD   cProvParam;
    #   PCRYPT_KEY_PROV_PARAM rgProvParam;
    ptr_size = ctypes.sizeof(ctypes.c_void_p)
    raw = buf.raw

    if len(raw) < ptr_size * 2 + 8:
        return None

    container_ptr = int.from_bytes(raw[0:ptr_size], "little")
    provider_ptr = int.from_bytes(raw[ptr_size : ptr_size * 2], "little")
    prov_type = int.from_bytes(raw[ptr_size * 2 : ptr_size * 2 + 4], "little")
    flags = int.from_bytes(raw[ptr_size * 2 + 4 : ptr_size * 2 + 8], "little")

    container_name = ""
    if container_ptr:
        try:
            container_name = ctypes.wstring_at(container_ptr) or ""
        except Exception:
            pass

    provider_name = ""
    if provider_ptr:
        try:
            provider_name = ctypes.wstring_at(provider_ptr) or ""
        except Exception:
            pass

    if not container_name and not provider_name:
        return None

    return {
        "ContainerName": container_name,
        "ProvName": provider_name,
        "ProvType": prov_type,
        "Flags": flags,
    }


def _test_key_accessible_silent(
    container_name: str,
    provider_name: str,
    provider_type: int,
    flags: int,
) -> bool:
    """Проверяет физическую доступность ключа БЕЗ UI-диалогов.

    Вызывает CryptAcquireContextW с флагом CRYPT_SILENT.
    Возвращает True, если ключевой контейнер можно открыть
    (токен вставлен или ключ в локальном хранилище).
    Возвращает False, если ключ недоступен (токен не вставлен).
    Никаких диалогов не показывается — CRYPT_SILENT подавляет UI CSP.
    """
    if not container_name or not provider_name:
        return False

    try:
        import ctypes

        advapi32 = ctypes.WinDLL("advapi32")

        acquire_flags = CRYPT_SILENT
        if flags & CRYPT_MACHINE_KEYSET:
            acquire_flags |= CRYPT_MACHINE_KEYSET

        hProv = ctypes.c_void_p()
        result = advapi32.CryptAcquireContextW(
            ctypes.byref(hProv),
            container_name,
            provider_name,
            provider_type,
            acquire_flags,
        )

        if result:
            advapi32.CryptReleaseContext(hProv, 0)
            return True

        return False
    except Exception:
        # Не удалось проверить (ctypes недоступен и т.п.) —
        # возвращаем True, чтобы избежать ложных отсечек.
        return True


# ═══════════════════════════════════════════════════════════════════════
# Парсинг сертификата
# ═══════════════════════════════════════════════════════════════════════

def _parse_cert(cert) -> dict | None:
    """Парсит COM-объект сертификата в словарь.

    Ключевое: НЕ обращается к cert.PrivateKey напрямую, чтобы
    избежать UI-диалогов КриптоПро для отключённых токенов.

    Алгоритм:
    1. Читаем базовые свойства (Subject, Issuer, Thumbprint, даты) —
       никогда не вызывают UI.
    2. Читаем CERT_KEY_PROV_INFO_PROP_ID через CryptoAPI —
       получаем ContainerName, ProvName, ProvType, Flags — БЕЗ UI.
    3. Вызываем CryptAcquireContextW с CRYPT_SILENT для проверки
       физической доступности ключа — БЕЗ UI.
    4. Если ключ доступен → безопасно читаем cert.PrivateKey.Unique-
       ContainerName (UI не появится, т.к. токен вставлен).
    5. Если ключ недоступен → has_private_key=False, сертификат
       будет отфильтрован.
    """
    try:
        # ── Базовые свойства (безопасные, без UI) ──────────────────────

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

        # ── Проверка закрытого ключа (SILENT) ──────────────────────────

        has_private_key = False
        container_name = ""

        # Шаг 1: Читаем провайдерную информацию БЕЗ обращения к COM PrivateKey
        prov_info = _read_cert_prov_info(cert)

        if prov_info:
            prov_container = prov_info.get("ContainerName", "")
            prov_name = prov_info.get("ProvName", "")
            prov_type = prov_info.get("ProvType", 0)
            prov_flags = prov_info.get("Flags", 0)

            _log(
                f"    🔑 ProvInfo: container={prov_container!r}, "
                f"provider={prov_name!r}, type={prov_type}, flags={prov_flags:#x}",
                "debug",
            )

            # Шаг 2: Тихо проверяем доступность ключевого контейнера
            accessible = _test_key_accessible_silent(
                prov_container, prov_name, prov_type, prov_flags
            )

            if accessible:
                # Шаг 3: Ключ доступен — безопасно читаем UniqueContainerName
                # (UI НЕ появится, т.к. токен физически вставлен)
                has_private_key = True
                try:
                    container_name = cert.PrivateKey.UniqueContainerName or ""
                except Exception:
                    try:
                        container_name = cert.PrivateKey.ContainerName or ""
                    except Exception:
                        # Fallback: используем логическое имя из prov_info
                        container_name = prov_container

                _log(
                    f"    ✅ Ключ доступен: {container_name!r}",
                    "debug",
                )
            else:
                # Ключ НЕ доступен (токен не вставлен) — пропускаем молча
                _log(
                    f"    🚫 Ключ недоступен (токен не вставлен): "
                    f"{prov_container!r} [{prov_name}]",
                    "debug",
                )
                has_private_key = False
                container_name = ""
        else:
            # Нет prov_info — cert.Handle недоступен или несовместим.
            # Fallback: проверяем HasPrivateKey() через COM, затем
            # пробуем PrivateKey (может показать UI для отключённых токенов,
            # но это единственный путь без CryptoAPI).
            _log("    ℹ ProvInfo недоступен, fallback к COM", "debug")

            try:
                has_private_key = bool(cert.HasPrivateKey())
            except Exception:
                try:
                    has_private_key = bool(cert.HasPrivateKey)
                except Exception:
                    pass

            if has_private_key:
                try:
                    container_name = cert.PrivateKey.UniqueContainerName or ""
                except Exception:
                    try:
                        container_name = cert.PrivateKey.ContainerName or ""
                    except Exception:
                        pass

        # ── ИНН из Subject ────────────────────────────────────────────
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
            "has_private_key": has_private_key,
            "container_name": container_name,
        }
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════════════
# Фильтрация и дедупликация
# ═══════════════════════════════════════════════════════════════════════

def list_certificates() -> list[dict]:
    """
    Возвращает список сертификатов УКЭП на съёмных носителях.

    Фильтрация:
    - Только действующие (не просроченные) сертификаты
    - Только сертификаты с закрытым ключом на съёмном носителе
      (USB-токен, RuToken, eToken, флешка)
    - Исключаются сертификаты из реестра и жёсткого диска
    - Дедупликация по thumbprint

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
        # ── Собираем из ВСЕХ доступных Store ────────────────────────
        all_certs: list[dict] = []
        sources: list[str] = []

        # 1. CAdESCOM.Store — КриптоПро CSP 5.x (основной)
        certs = _list_certs_cadescom_store()
        if certs:
            all_certs.extend(certs)
            sources.append("CAdESCOM.Store")

        # 2. CAPICOM.Store — Windows Certificate Store
        certs = _list_certs_capicom_store()
        if certs:
            all_certs.extend(certs)
            sources.append("CAPICOM.Store")

        # 3. CPCSPStore — legacy CSP 4.x
        certs = _list_certs_legacy_store()
        if certs:
            all_certs.extend(certs)
            sources.append("CPCSPStore")

        if not all_certs:
            _log(
                "⚠ Подходящие сертификаты не найдены.\n"
                "   Проверьте:\n"
                "   • USB-токен (RuToken/eToken) подключён к ПК\n"
                "   • Сертификат на токене не просрочен\n"
                "   • КриптоПро CSP установлен и лицензия активна\n"
                "   • Сертификат виден в КриптоПро CSP → Сервис → "
                "Просмотреть сертификаты",
                "warn",
            )
            return []

        _log(f"  📋 Всего сырых сертификатов из {sources}: {len(all_certs)}")

        # ── Дедупликация по thumbprint ──────────────────────────────
        seen: dict[str, bool] = {}
        unique_certs: list[dict] = []
        for cert in all_certs:
            tp = cert.get("thumbprint", "").strip().lower().replace(" ", "")
            if tp and tp not in seen:
                seen[tp] = True
                unique_certs.append(cert)

        dupes = len(all_certs) - len(unique_certs)
        if dupes > 0:
            _log(f"  🔄 Дедупликация: убрано {dupes} дубликатов")

        # ── Фильтрация ──────────────────────────────────────────────
        filtered = _filter_certs(unique_certs)

        _log(
            f"📋 Уникальных: {len(unique_certs)}, "
            f"на съёмных носителях (действующих): {len(filtered)}",
            "success",
        )

        return filtered


def _filter_certs(certs: list[dict]) -> list[dict]:
    """Фильтрует сертификаты: только действующие на съёмных носителях.

    Важно: сертификаты с has_private_key=False уже отсеяны в _parse_cert
    (ключ недоступен → has_private_key=False). Поэтому здесь мы только
    проверяем сроки и тип носителя.
    """
    result = []
    for cert in certs:
        # Пропускаем просроченные
        if _is_expired(cert.get("not_after", "")):
            _log(
                f"  ⏭ Пропущен (просрочен): {cert.get('subject', '?')[:60]}..."
                f" (до {cert.get('not_after', '?')})"
            )
            continue

        # Пропускаем без закрытого ключа
        # (ключ недоступен — токен не вставлен, обработано в _parse_cert)
        if not cert.get("has_private_key", False):
            _log(
                f"  ⏭ Пропущен (нет закрытого ключа): "
                f"{cert.get('subject', '?')[:60]}..."
            )
            continue

        # Пропускаем ключи из реестра/диска (не на съёмном носителе)
        container = cert.get("container_name", "")
        if container and not _is_on_removable_media(container):
            _log(
                f"  ⏭ Пропущен (локальное хранилище): "
                f"{cert.get('subject', '?')[:60]}... [{container}]",
            )
            continue

        # Если container_name пуст, но has_private_key=True —
        # ключ доступен, но мы не знаем точный путь. Не отсеиваем,
        # потому что _test_key_accessible_silent уже подтвердила
        # доступность (токен вставлен), а _is_on_removable_media
        # не может дать False для пустой строки.
        if not container:
            _log(
                f"  ℹ Нет имени контейнера, но ключ доступен: "
                f"{cert.get('subject', '?')[:60]}...",
                "debug",
            )

        result.append(cert)
    return result


# ═══════════════════════════════════════════════════════════════════════
# Упрощённый список сертификатов (без обращения к PrivateKey / CryptoAPI)
# ═══════════════════════════════════════════════════════════════════════

def _parse_cert_simple(cert) -> dict | None:
    """Парсит COM-объект сертификата в словарь — БЕЗОПАСНЫЙ вариант.

    НЕ обращается к cert.PrivateKey, cert.Handle, CryptoAPI — НИКОГДА
    не вызывает UI-диалоги КриптоПро. Читает только:
    - SubjectName / Subject
    - IssuerName / Issuer
    - Thumbprint
    - ValidFromDate / ValidToDate
    - ИНН из Subject DN
    """
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

        # ИНН из Subject
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


def _list_certs_simple_from_store(prog_id: str) -> list[dict]:
    """Перечисляет сертификаты из COM-хранилища БЕЗ обращения к PrivateKey.

    Использует _parse_cert_simple() — никаких UI-диалогов КриптоПро.
    """
    try:
        import win32com.client
        store = win32com.client.Dispatch(prog_id)
    except Exception as e:
        _log(f"  ℹ {prog_id} недоступен: {e}")
        return []

    certs = []
    try:
        store.Open(
            CAPICOM_CURRENT_USER_STORE,
            "My",
            CAPICOM_STORE_OPEN_READ_ONLY,
        )

        count = store.Certificates.Count
        _log(f"  📂 {prog_id}: {count} сертификатов")

        for i in range(1, count + 1):
            try:
                cert = store.Certificates.Item(i)
                info = _parse_cert_simple(cert)
                if info:
                    certs.append(info)
            except Exception as e:
                _log(f"  ⚠ Ошибка сертификата #{i} ({prog_id}): {e}")

        store.Close()
    except Exception as e:
        _log(f"  ⚠ {prog_id} ошибка: {e}")
        try:
            store.Close()
        except Exception:
            pass

    return certs


def list_all_valid_certificates() -> list[dict]:
    """Возвращает все действующие (не просроченные) сертификаты.

    В отличие от list_certificates():
    - НЕ проверяет наличие закрытого ключа (нет обращения к PrivateKey)
    - НЕ фильтрует по типу носителя (съёмный / реестр)
    - НЕ вызывает CryptoAPI → НИКОГДА не показывает UI-диалоги КриптоПро

    Фильтрация:
    - Только действующие (не просроченные) по ValidToDate
    - Дедупликация по thumbprint

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
        all_certs: list[dict] = []
        sources: list[str] = []

        for prog_id in ("CAdESCOM.Store", "CAPICOM.Store", "CPCSPStore.Store"):
            certs = _list_certs_simple_from_store(prog_id)
            if certs:
                all_certs.extend(certs)
                sources.append(prog_id)

        if not all_certs:
            _log(
                "⚠ Сертификаты не найдены.\n"
                "   Проверьте:\n"
                "   • КриптоПро CSP установлен и лицензия активна\n"
                "   • Сертификат установлен в хранилище «Личные»\n"
                "   • Сертификат виден в КриптоПро CSP → Сервис → "
                "Просмотреть сертификаты",
                "warn",
            )
            return []

        _log(f"  📋 Всего сертификатов из {sources}: {len(all_certs)}")

        # Дедупликация по thumbprint
        seen: dict[str, bool] = {}
        unique: list[dict] = []
        for cert in all_certs:
            tp = cert.get("thumbprint", "").strip().lower().replace(" ", "")
            if tp and tp not in seen:
                seen[tp] = True
                unique.append(cert)

        dupes = len(all_certs) - len(unique)
        if dupes > 0:
            _log(f"  🔄 Дедупликация: убрано {dupes} дубликатов")

        # Фильтрация: только действующие (не просроченные)
        valid: list[dict] = []
        for cert in unique:
            if _is_expired(cert.get("not_after", "")):
                _log(
                    f"  ⏭ Пропущен (просрочен): "
                    f"{cert.get('subject', '?')[:60]}..."
                )
                continue
            valid.append(cert)

        _log(
            f"📋 Уникальных: {len(unique)}, действующих: {len(valid)}",
            "success",
        )

        return valid


# ═══════════════════════════════════════════════════════════════════════
# Перечисление сертификатов из COM-хранилищ
# ═══════════════════════════════════════════════════════════════════════

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


# ═══════════════════════════════════════════════════════════════════════
# Вспомогательные функции
# ═══════════════════════════════════════════════════════════════════════

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
                            inn_val = bytes.fromhex(inn_hex).decode(
                                "ascii", errors="ignore"
                            ).strip()
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

        for prog_id in [
            "CAdESCOM.Store", "CAPICOM.Store", "CPCSPStore.Store",
            "CAdESCOM.CPSigner", "CAdESCOM.CadesSignedData",
            "CAdESCOM.About",
        ]:
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