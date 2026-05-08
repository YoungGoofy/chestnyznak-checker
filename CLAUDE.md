# CLAUDE.md — ChestnyZnakChecker

Проект для проверки кодов маркировки системы «Честный Знак» (CRPT) через публичный и закрытый (True) API с выгрузкой в Excel.

## Структура

```
chestnyznak_checker/
├── check_codes.py        # CLI-скрипт: вся бизнес-логика, HTTP, парсинг, Excel
├── gui_app.py            # GUI (tkinter): импортирует функции из check_codes.py
├── requirements.txt      # python-dotenv, openpyxl
├── .env                  # CHESTNYZNAK_TOKEN=... (gitignore)
├── .token_timestamp      # создаётся GUI, хранит unixtime последнего сохранения токена
├── flake.nix             # nix-shell для NixOS (tkinter + openpyxl)
├── shell.nix             # nix-shell без flakes
├── build_instructions.txt # pyinstaller для сборки в .exe
├── README.md             # инструкция для пользователя
└── CLAUDE.md             # этот файл
```

## Ключевые компоненты check_codes.py

### Эндпоинты
- Публичный: `POST https://mobile.api.crpt.ru/mobile/check` — не отдаёт владельца
- True API: `POST https://markirovka.crpt.ru/api/v3/true-api/cises/info?pg=<group>` — нужен токен

### Структура ответа True API
```json
[{"cisInfo": {"requestedCis": "...", "ownerName": "...", "ownerInn": "...", "producerName": "...", "introducedDate": "ISO", "status": "...", ...}}]
```
Все данные ЛЕЖАТ ВНУТРИ `cisInfo` — это важно для парсинга.

### Функции (все импортируются в gui_app.py)
- `load_env(script_dir)` — загружает `.env`
- `public_check(code)` → dict | None
- `get_pg_from_public(data)` → str (категория из публичного ответа)
- `true_check_batch(codes, pg, token)` → list[dict] | None
- `true_check_with_retry_pg(codes, category, token, debug=False, log_fn=None)` → tuple[int|None, list[dict]|None] — пробуетpg по CATEGORY_TO_PG
- `true_check_auto(codes, token)` → {code: item} — определяет pg через публичный API, потом True API (в GUI не используется)
- `parse_result(code, data, mode)` → list[str] (10 колонок Excel)
- `save_excel(rows, output_path)` — пишет .xlsx с Catppuccin-стилем
- `http_post(url, payload, headers)` — низкоуровневый POST с ретраями

### Формат строки Excel (ровно 10 колонок)
```
[Штрихкод, GTIN, Бренд, Индекс картинки, Статус, 1, Владелец, Производитель, Дата ввода, Способ ввода]
```

### Маппинги
- `CATEGORY_TO_PG` — `lp→lp`, `milk→milk` и т.д.
- `PG_ALIASES` — варианты pg для перебора при 404
- `STATUS_MAP` — EMITTED→Эмитирован, INTRODUCED→В обороте...
- `EMISSION_TYPE_MAP` — REMAINS→Маркировка остатков, REMARK→Перемаркировка...

## GUI (gui_app.py)

- **Фреймворк**: tkinter (встроен в Python, кроссплатформен)
- **Потоковая модель**: проверка в `threading.Thread(daemon=True)`, логи через `queue.Queue` → `root.after(100, poll)`
- **Проверка токена**: `_quick_auth_check(token, pg_code)` делает тестовый запрос перед запуском
- **Товарные группы**: выпадающий список `PRODUCT_GROUPS` — 15 ТГ с понятными русскими названиями. Дефолт: «Лёгкая промышленность». Выбранная ТГ передаётся в `true_check_with_retry_pg()`.
- **Батчевая обработка**: коды отправляются пачками по `BATCH_SIZE=100`, прогресс отображается в логах.
- **Остановка**: `_stop_requested` флаг проверяется между батчами.
- **Ошибки API**: `explain_api_error()` расшифровывает HTTP-статусы (401, 403, 429, 451, 5xx) с пояснениями для пользователя.
- **Неверная ТГ**: если ни один код из батча не распознан, выводится предупреждение «Возможно выбрана неверная товарная группа».
- **Таймстамп токена**: `.token_timestamp` в папке скрипта, проверка ≥8 часов при старте
- **Цвета**: Catppuccin Mocha

### Импорт из check_codes.py
```python
from check_codes import (
    public_check, true_check_batch, true_check_with_retry_pg,
    get_pg_from_public, parse_result, save_excel,
    EXCEL_HEADERS, load_env,
)
```

## Зависимости

| Пакет | Для чего |
|-------|----------|
| `openpyxl` | Excel-экспорт |
| `python-dotenv` | .env загрузка (опционально, load_env работает и без него) |
| `tkinter` | GUI (встроен в Python, на NixOS через `nix-shell -p tk`) |

## Запуск

```bash
# CLI
python check_codes.py --true -f codes.txt -o result.xlsx

# GUI
python gui_app.py

# NixOS GUI
nix-shell -p python311 -p tk --run "source .venv/bin/activate && python gui_app.py"
```

## Важные нюансы

1. **Геоблокировка**: API доступен только с российских IP. Тесты с сервера вне РФ будут падать.
2. **Токен живёт 8–12 часов**, берётся из `localStorage.getItem('token')` в консоли ЛК.
3. **Даты в True API — ISO-строки**, в публичном — миллисекундные timestamp. Функция `iso_to_datetime_str` и `ts_to_datetime_str` обрабатывают оба.
4. **Коды содержат спецсимволы** (`!`, `*`, `'`, `)`, `"`, `<`, `>` и т.д.) — в CLI лучше передавать через файл, не через аргументы командной строки.
5. **404 в True API** — может означать как «код не найден», так и «неверная товарная группа». Скрипт перебирает варианты pg через `PG_ALIASES`.
6. **Excel создаётся даже без openpyxl** — fallback в CSV.
7. **GUI не блокируется** — вся работа в фоновом потоке, логи через очередь.

## Сборка в .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --hidden-import openpyxl --name ChestnyZnakChecker gui_app.py
```

**Важно:** `get_app_dir()` использует `sys.executable` в .exe (а не `__file__`), поэтому `.env` и `.token_timestamp` создаются **в папке рядом с .exe**, а не во временной. При передаче .exe другому пользователю:
- `.env` создастся автоматически при первом сохранении токена через GUI
- `.token_timestamp` обновится при сохранении токена
- Можно положить готовый `.env` рядом с .exe — он подхватится

## Если что-то сломалось

- Проверить токен: `curl -X POST "https://markirovka.crpt.ru/api/v3/true-api/cises/info?pg=lp" -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" -d '["любой_код"]'`
- Проверить публичный API: `curl -X POST "https://mobile.api.crpt.ru/mobile/check" -H "Content-Type: application/json" -d '{"code":"любой_код"}'`
- Запустить CLI с `--debug`: покажет полные URL и ответы
- GUI логирует всё в цветное текстовое поле, ошибки выделены красным
