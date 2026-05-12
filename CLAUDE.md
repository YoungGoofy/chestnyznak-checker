# CLAUDE.md — CISChecker

Проект для проверки кодов маркировки системы «Честный Знак» (CRPT) через публичный и закрытый (True) API с выгрузкой в Excel.

**Репозиторий:** https://github.com/YoungGoofy/chestnyznak-checker
**Текущая версия:** `APP_VERSION = "1.2.1"` (в gui_app.py)

## Структура

```
chestnyznak_checker/
├── check_codes.py            # CLI + бизнес-логика: HTTP, парсинг, Excel
├── crypto_auth.py             # Авторизация через УКЭП (КриптоПро CSP)
├── gui_app.py                 # GUI (tkinter): импортирует check_codes + updater + crypto_auth
├── updater.py                 # Автообновление через GitHub Releases
├── .github/
│   └── workflows/
│       └── build.yml          # GitHub Actions: автосборка .exe при пуше тега v*
├── requirements.txt           # openpyxl, python-dotenv
├── build.bat                  # Сборка .exe на Windows (PyInstaller)
├── setup.bat                  # Создание .venv + установка зависимостей
├── run.bat                    # Запуск GUI из .venv на Windows
├── run_nix.sh                 # Запуск GUI на NixOS
├── flake.nix                  # nix-shell для NixOS (flakes)
├── shell.nix                  # nix-shell без flakes
├── .gitignore                 # исключает .env, __pycache__, .venv, dist, build
├── README.md                  # Инструкция для пользователя
└── CLAUDE.md                  # этот файл
```

**Не в репозитории** (gitignore):
- `.env` — токен ChZ, ИНН (секрет)
- `.token_timestamp` — unixtime последнего сохранения токена
- `__pycache__/`, `.venv/`, `build/`, `dist/`, `*.spec`

## Ключевые компоненты check_codes.py

### Эндпоинты
- Публичный: `POST https://mobile.api.crpt.ru/mobile/check` — не отдаёт владельца
- True API: `POST https://markirovka.crpt.ru/api/v3/true-api/cises/info?pg=<group>` — нужен токен

### Структура ответа True API
```json
[{"cisInfo": {"requestedCis": "...", "gtin": "...", "brand": "...", "ownerName": "...", "ownerInn": "...", "producerName": "...", "introducedDate": "ISO", "status": "...", ...}}]
```
Все данные ЛЕЖАТ ВНУТРИ `cisInfo` — это важно для парсинга.

### Функции (все импортируются в gui_app.py)
- `load_env(script_dir)` — загружает `.env`
- `public_check(code)` → dict | None
- `get_pg_from_public(data)` → str (категория из публичного ответа)
- `true_check_batch(codes, pg, token, log_fn=None)` → **tuple[int|None, list[dict]|None]** — возвращает (http_status, results). Статус None = сетевая ошибка, results None = все коды не найдены.
- `true_check_with_retry_pg(codes, category, token, debug=False, log_fn=None)` → tuple[int|None, list[dict]|None] — пробует pg по CATEGORY_TO_PG, пропускает другие pg при 401/403/429/451
- `true_check_auto(codes, token, log_fn=None, stop_fn=None, batch_size=100, progress_fn=None)` → dict — батчевая обработка с автоопределением ТГ через публичный API (в GUI **не используется**, ТГ выбирается вручную)
- `parse_result(code, data, mode)` → list[str] (10 колонок Excel)
- `parse_public_row(data)` → list[str] (10 колонок из публичного ответа)
- `save_excel(rows, output_path)` — пишет .xlsx с Catppuccin-стилем
- `http_post(url, payload, headers)` — низкоуровневый POST с ретраями
- `explain_http_status(status_code)` → str — расшифровка HTTP-ошибок для пользователя

### Формат строки Excel (ровно 10 колонок)
```
[Штрихкод, GTIN, Бренд, Индекс картинки, Статус, Количество, Владелец, Производитель, Дата ввода, Способ ввода]
```
Индексы: 0=Штрихкод, 1=GTIN, 2=Бренд, 3=Индекс картинки, 4=Статус, 5=Количество, 6=Владелец, 7=Производитель, 8=Дата ввода, 9=Способ ввода

### Источники данных для GTIN и Бренд
**True API:**
- GTIN → `cisInfo.gtin`
- Бренд → `cisInfo.brand`

**Публичный API:**
- GTIN → `codeResolveData.gtin` или `codeData.gtin`
- Бренд → `catalogData[0].brand_name` или `group.brand`

### Маппинги
- `CATEGORY_TO_PG` — `lp→lp`, `milk→milk` и т.д.
- `PG_ALIASES` — варианты pg для перебора при 404
- `STATUS_MAP` — EMITTED→Эмитирован, INTRODUCED→В обороте...
- `EMISSION_TYPE_MAP` — REMAINS→Маркировка остатков, REMARK→Перемаркировка...

## GUI (gui_app.py)

### Константы
- `APP_VERSION = "1.1"` — версия для проверки обновлений
- `BATCH_SIZE = 100` — размер батча для True API
- `PRODUCT_GROUPS` — словарь {русское название: код pg}, 15 товарных групп

### Товарные группы (выпадающий список)
Дефолт: «Лёгкая промышленность (одежда, бельё, текстиль)» → `lp`

Словарь `PRODUCT_GROUPS`:
- Лёгкая промышленность → `lp`
- Молочная продукция → `milk`
- Табак → `tobacco`
- Упакованная вода → `water`
- Пиво и пивные напитки → `beer`
- Обувь → `shoes`
- Парфюмерия → `perfume`
- Шины → `tires`
- Фото- и видеокамеры → `camera`
- Велосипеды → `bicycle`
- Меховые изделия → `furs`
- Лекарственные средства → `medicine`
- БАД → `bio`
- Дезинфицирующие средства → `antiseptic`
- Кресла-коляски → `wheelchair`

### Архитектура
- **Фреймворк**: tkinter (встроен в Python, кроссплатформен)
- **Потоковая модель**: проверка в `threading.Thread(daemon=True)`, логи через `queue.Queue` → `root.after(100, poll)`
- **Проверка токена**: `_quick_auth_check(token, pg_code)` делает тестовый запрос с выбранной ТГ перед запуском
- **Батчевая обработка**: коды отправляются пачками по `BATCH_SIZE=100` через `true_check_with_retry_pg()`, прогресс `[N/total]` отображается в логах
- **Остановка**: `_stop_requested` флаг проверяется между батчами
- **Ошибки API**: `explain_api_error()` расшифровывает HTTP-статусы (401, 403, 429, 451, 5xx) с пояснениями для пользователя
- **Неверная ТГ**: если ни один код из батча не распознан, выводится предупреждение «Возможно выбрана неверная товарная группа»
- **Таймстамп токена**: `.token_timestamp` в папке скрипта, проверка ≥8 часов при старте
- **Цвета**: Catppuccin Mocha

### Меню
- **Файл**: Загрузить коды из файла, Выход
- **Настройки**: Токен True API...
- **Справка**: Инструкция, Проверить обновления..., О программе

### Кнопка-индикатор обновлений
- Расположена в верхней панели (справа от статуса токена)
- Цветовая схема Catppuccin:
  - `🔄 Проверка...` (серый, DISABLED) — идёт проверка
  - `✓ Обновлений нет` (серый, DISABLED) — текущая версия актуальна
  - `⬆ Обновление v1.2` (зелёный, NORMAL) — доступно обновление, кликабельна
- При нажатии на зелёную кнопку — диалог подтверждения, затем скачивание
- **Автопроверка при старте**: через 1.5 сек после запуска (тихая, без диалогов)
- **Ручная проверка**: через меню Справка → Проверить обновления (с диалогом)
- Сохраняет `_pending_update` для повторного нажатия

### Импорты
```python
from check_codes import (
    public_check, true_check_batch, true_check_with_retry_pg,
    get_pg_from_public, parse_result, save_excel,
    EXCEL_HEADERS, load_env,
)
from crypto_auth import (
    auth_jwt, list_certificates,
    set_log_fn as set_auth_log_fn,
)
from updater import (
    check_for_update as _check_update,
    perform_update as _perform_update,
    is_frozen as _is_frozen,
    GITHUB_REPO,
)
```

### Ключевые методы класса App
- `_start_true_thread()` — запуск проверки через True API с выбранной ТГ
- `_start_public_thread()` — запуск проверки через публичный API (fallback)
- `_stop_processing()` — установка флага `_stop_requested`
- `_quick_auth_check(token, pg_code)` — тестовый запрос для проверки токена
- `_open_token_dialog()` — ручной ввод токена
- `_open_ukep_dialog()` — авторизация через УКЭП (сертификат + ИНН → токен)
- `_auto_check_updates()` — автопроверка обновлений при старте (тихая)
- `_run_update_check()` — ручная проверка обновлений из меню (с диалогом)
- `_on_update_button()` — обработчик нажатия кнопки-индикатора обновлений
- `_check_updates()` — делегирует в `_run_update_check()`
- `_apply_update(release_info)` — скачивание и подмена .exe

## Модуль авторизации через УКЭП (crypto_auth.py)

### Метод авторизации
**JWT flow** (единственный метод) — два шага: GET `/true-api/auth/key` → подпись challenge → POST `/true-api/auth/simpleSignIn`

United Token был **удалён в v1.2.0** — остался только JWT flow.

### Ключевые функции
- `auth_jwt(thumbprint="")` → `(bool, str)` — JWT авторизация (2 шага)
- `list_certificates()` → `list[dict]` — список сертификатов УКЭП (только действующие, только на съёмных носителях — USB-токен/флешка)
- `sign_data(data, thumbprint="")` → `bytes | None` — подпись данных УКЭП (attached CMS)
- `set_log_fn(fn)` — подключить функцию логирования (для GUI)

### Порядок поиска сертификатов
1. **CAPICOM.Store** (Windows Certificate Store — перехватывается КриптоПро CSP 5.x/4.x)
2. **CPCSPStore.Store** (CSP 4.x — legacy)
3. **cryptcp CLI** (fallback для Linux)

### Порядок подписи
1. **CAdESCOM.Store/CAPICOM.Store + CAdESCOM.CPSigner + CAdESCOM.CadesSignedData.SignCades(CADES_BES)** — Windows COM (CSP 5.x)
2. **CPCSPStore + CPSigner + CPSignedData** — legacy CSP 4.x
3. **cryptcp CLI** — fallback

### Подпись: важные детали
- **SignCades()** а НЕ `Sign()` — `Sign()` создаёт CAdES-X Long Type 1 и требует TSP-сервер → падает.
- **ContentEncoding = 1** (CADESCOM_BASE64_TO_BINARY) обязателен ПЕРЕД установкой Content.
- **Подпись attached** (присоединённая) — требуется для simpleSignIn.
- **Переносы строк** (\r\n) удаляются из base64-результата — API ЧЗ их не принимает.

### Важные детали
- **CAdESCOM.Store** — ProgID, предоставляемый КриптоПро CSP 5.x для доступа к хранилищу сертификатов (основной метод). **CAPICOM.Store** — стандартный Windows COM ProgID (fallback).
- **CAdESCOM.CPSigner** и **CAdESCOM.CadesSignedData** — правильные ProgID для подписи (КриптоПро CSP 5.x)
- **CPCSPStore.Store** — legacy ProgID для CSP 4.x, требует `Open(StoreLocation, StoreName, OpenMode)`
- **pywin32** нужен для COM-доступа на Windows (`pip install pywin32`), добавлен в `requirements.txt` с условием `sys_platform == "win32"`. Включает `pythoncom` для COM-инициализации.
- **COM и потоки**: все COM-вызовы обёрнуты в `_com_initialized()` (contextmanager) для `pythoncom.CoInitialize()/CoUninitialize()` — это **КРИТИЧЕСКИ ВАЖНО** для фоновых потоков (threading.Thread).
- Подпись **присоединённая** (attached CMS), кодируется в base64
- ИНН для JWT flow — извлекается из сертификата (OID `1.2.643.3.131.1.1` или Subject DN)

### Диалог УКЭП в GUI
- Меню: Настройки → «🔐 Получить токен через УКЭП»
- Список сертификатов (Listbox), метод: JWT через УКЭП
- Кнопка «🔄 Обновить список» — перечитывает сертификаты
- Кнопка «🔐 Получить токен» → фоновый поток → `auth_jwt(thumbprint)` → сохранение в .env
- При сохранении токена пишется файл `.token_expires` с unix timestamp истечения (из JWT `exp` payload)

## Модуль обновлений (updater.py)

### Константы
- `GITHUB_REPO = "YoungGoofy/chestnyznak-checker"`
- `RELEASES_API = "https://api.github.com/repos/.../releases/latest"`
- `EXE_NAME = "CISChecker.exe"`

### Функции
- `get_current_version()` → str — возвращает `APP_VERSION` из gui_app
- `is_frozen()` → bool — True, если запущен как PyInstaller .exe
- `get_exe_path()` → Path — путь к текущему .exe/.py
- `fetch_latest_release()` → dict | None — информация о последнем релизе с GitHub
- `compare_versions(current, latest)` → int — 1=есть обновление, 0=равны, -1=текущая новее
- `download_exe(url, dest, progress_fn)` → bool — скачивание .exe с прогрессом
- `create_updater_bat(old_exe, new_exe)` → Path — .bat скрипт для подмены .exe
- `perform_update(exe_url, progress_fn)` → tuple[bool, str] — полный цикл обновления
- `check_for_update()` → tuple[bool, str, dict|None] — проверка наличия обновления

### Как работает автообновление

1. Пользователь нажимает **Справка → Проверить обновления**
2. Приложение стучится в GitHub Releases API
3. Сравнивает `tag_name` (напр. `v1.2`) с текущей `APP_VERSION`
4. Если доступна новая версия:
   - **В .exe-режиме**: скачивает новый `.exe` в `_update_tmp/` → создаёт `_update.bat` → запускает батник → закрывает приложение → батник ждёт завершения процесса → подменяет .exe → запускает новый → удаляет .bat
   - **В .py-режиме**: открывает браузер на страницу релиза GitHub

## Зависимости

- `openpyxl` — Excel-экспорт (обязательно)
- `python-dotenv` — .env загрузка (опционально, load_env работает и без него)
- `tkinter` — GUI (встроен в Python, на NixOS через `nix-shell -p tk`)

## Запуск

```bash
# CLI
python check_codes.py --true -f codes.txt -o result.xlsx

# GUI (Python)
python gui_app.py

# GUI (Windows .exe)
CISChecker.exe

# NixOS GUI
nix-shell -p python311 -p tk --run "source .venv/bin/activate && python gui_app.py"
```

## Сборка в .exe

```bash
# Вручную
pip install pyinstaller
pyinstaller --onefile --windowed --hidden-import openpyxl --name CISChecker gui_app.py

# Через build.bat (Windows)
build.bat
```

**Важно:** `get_app_dir()` использует `sys.executable` в .exe (а не `__file__`), поэтому `.env` и `.token_timestamp` создаются **в папке рядом с .exe**, а не во временной. При передаче .exe другому пользователю:
- `.env` создастся автоматически при первом сохранении токена через GUI
- `.token_timestamp` обновится при сохранении токена
- Можно положить готовый `.env` рядом с .exe — он подхватится

## GitHub Actions — автосборка .exe

**Файл:** `.github/workflows/build.yml`

Срабатывает при пуше тега `v*` (напр. `v1.0`, `v1.1`):
1. Запускается на `windows-latest`
2. Ставит Python 3.11 + зависимости
3. Собирает `CISChecker.exe` через PyInstaller
4. Создаёт GitHub Release с `.exe` файлом

### Как выпустить обновление

1. Обновить `APP_VERSION` в `gui_app.py` (напр. `"1.1"`)
2. `git add -A && git commit -m "v1.1"`
3. `git tag v1.1`
4. `git push origin main --tags`
5. GitHub Actions **автоматически** соберёт `.exe` и создаст Release

**Для пуша нужен GitHub PAT с правами `repo` + `workflow`** (право workflow обязательно для `.github/workflows/`).

## Важные нюансы

1. **Геоблокировка**: API доступен только с российских IP. Тесты с сервера вне РФ будут падать.
2. **Токен живёт 8–12 часов**, берётся из `localStorage.getItem('token')` в консоли ЛК Честного Знака.
3. **Даты в True API — ISO-строки**, в публичном — миллисекундные timestamp. Функция `iso_to_datetime_str` и `ts_to_datetime_str` обрабатывают оба.
4. **Коды содержат спецсимволы** (`!`, `*`, `'`, `)`, `\"`, `<`, `>` и т.д.) — в CLI лучше передавать через файл, не через аргументы командной строки.
5. **404 в True API** — может означать как «код не найден», так и «неверная товарная группа». Скрипт перебирает варианты pg через `PG_ALIASES`.
6. **Excel создаётся даже без openpyxl** — fallback в CSV.
7. **GUI не блокируется** — вся работа в фоновом потоке, логи через очередь.
8. **true_check_batch** возвращает tuple `(http_status, results)`, а не просто results. Это важное отличие — upstream код должен проверять оба значения.
9. **Ошибка строк в Excel** — ровно 10 колонок, пустые строки для GTIN/Бренд/Индекс картинки при ошибках.
10. **Индексы колонок**: Статус = `[4]` (было `[2]`), Владелец = `[6]` (было `[4]`) — после добавления GTIN и Бренд.

## Если что-то сломалось

- Проверить токен: `curl -X POST "https://markirovka.crpt.ru/api/v3/true-api/cises/info?pg=lp" -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" -d '["любой_код"]'`
- Проверить публичный API: `curl -X POST "https://mobile.api.crpt.ru/mobile/check" -H "Content-Type: application/json" -d '{"code":"любой_код"}'`
- Запустить CLI с `--debug`: покажет полные URL и ответы
- GUI логирует всё в цветное текстовое поле, ошибки выделены красным
- Проверить GitHub Release: https://github.com/YoungGoofy/chestnyznak-checker/releases
- Проверить CI: https://github.com/YoungGoofy/chestnyznak-checker/actions