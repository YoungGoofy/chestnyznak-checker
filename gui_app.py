#!/usr/bin/env python3
"""
GUI для проверки кодов маркировки (CISChecker).
Tkinter — работает на Windows, Linux, NixOS без внешних зависимостей.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import queue
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.error import URLError
from tkinter import (
    Tk, Frame, Label, Button, Text, Scrollbar, Menu,
    Toplevel, Entry, filedialog, messagebox,
    END, N, S, E, W, VERTICAL, HORIZONTAL, DISABLED, NORMAL,
    StringVar,
)
from tkinter import ttk

# ══════════════════════════════════════════════════════════════════════
# Определение папки приложения (работает и в .py, и в .exe)
# ══════════════════════════════════════════════════════════════════════

def get_app_dir() -> Path:
    """Папка, где лежит запущенный скрипт/.exe.
    В скомпилированном PyInstaller .exe: sys.executable → папка рядом с exe.
    В обычном Python: __file__ → папка скрипта."""
    if getattr(sys, 'frozen', False):
        # PyInstaller — .exe
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


SCRIPT_DIR = get_app_dir()
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

# Импортируем нужные функции из check_codes
from check_codes import (
    public_check,
    true_check_batch,
    true_check_with_retry_pg,
    get_pg_from_public,
    parse_result,
    save_excel,
    EXCEL_HEADERS,
    load_env,
)

# Импортируем модуль авторизации через УКЭП
from crypto_auth import (
    auth_jwt,
    list_certificates,
    set_log_fn as set_auth_log_fn,
)

# Импортируем модуль обновления
from updater import (
    check_for_update as _check_update,
    perform_update as _perform_update,
    is_frozen as _is_frozen,
    GITHUB_REPO,
)

# ══════════════════════════════════════════════════════════════════════
# Константы
# ══════════════════════════════════════════════════════════════════════
APP_TITLE = "CISChecker — Проверка кодов маркировки"
APP_VERSION = "1.2.1"
WINDOW_WIDTH = 850
WINDOW_HEIGHT = 620
LOG_MAX_LINES = 500
BATCH_SIZE = 100

# Товарные группы: русское название → код pg для API
PRODUCT_GROUPS: dict[str, str] = {
    "Лёгкая промышленность (одежда, бельё, текстиль)": "lp",
    "Молочная продукция": "milk",
    "Табак": "tobacco",
    "Упакованная вода": "water",
    "Пиво и пивные напитки": "beer",
    "Обувь": "shoes",
    "Парфюмерия": "perfume",
    "Шины": "tires",
    "Фото- и видеокамеры": "camera",
    "Велосипеды": "bicycle",
    "Меховые изделия": "furs",
    "Лекарственные средства": "medicine",
    "БАД (биологически активные добавки)": "bio",
    "Дезинфицирующие средства (антисептики)": "antiseptic",
    "Кресла-коляски": "wheelchair",
}
PRODUCT_GROUPS_DEFAULT = "Лёгкая промышленность (одежда, бельё, текстиль)"

# Цвета
COLOR_BG = "#1e1e2e"
COLOR_FRAME_BG = "#181825"
COLOR_BUTTON_BG = "#45475a"
COLOR_BUTTON_FG = "#cdd6f4"
COLOR_BUTTON_ACTIVE_BG = "#585b70"
COLOR_BUTTON_ACCENT = "#89b4fa"
COLOR_BUTTON_ACCENT_HOVER = "#74c7ec"
COLOR_LOG_BG = "#11111b"
COLOR_LOG_FG = "#cdd6f4"
COLOR_LOG_ERROR = "#f38ba8"
COLOR_LOG_SUCCESS = "#a6e3a1"
COLOR_LOG_INFO = "#89b4fa"
COLOR_LOG_WARN = "#fab387"
COLOR_HEADER_BG = "#313244"
COLOR_UPDATE_NONE = "#585b70"      # серый — обновлений нет
COLOR_UPDATE_AVAIL = "#a6e3a1"     # зелёный — доступно обновление
COLOR_UPDATE_FG_NONE = "#cdd6f4"   # текст на серой кнопке
COLOR_UPDATE_FG_AVAIL = "#1e1e2e"  # тёмный текст на зелёной кнопке

# ══════════════════════════════════════════════════════════════════════
# Очередь для передачи логов из фонового потока в GUI
# ══════════════════════════════════════════════════════════════════════
log_queue: queue.Queue = queue.Queue()


def log_to_gui(message: str, tag: str = "info") -> None:
    """Кладёт сообщение в очередь для отображения в GUI."""
    log_queue.put((message, tag))


def explain_api_error(status: int | None, body: str | None) -> str:
    """Возвращает человекочитаемое описание ошибки API."""
    if status is None:
        return "Не удалось подключиться к серверу. Проверьте интернет-соединение и убедитесь, что IP российский (API геоблокирует зарубежные IP)."
    messages = {
        400: "Некорректный запрос (HTTP 400). Сервер не смог обработать отправленные данные.",
        401: "Токен недействителен или просрочен (HTTP 401). Получите новый токен в ЛК Честного Знака.",
        403: "Доступ запрещён (HTTP 403). У вашего токена нет прав для данной операции.",
        404: "Код не найден или неверная товарная группа (HTTP 404).",
        429: "Слишком много запросов (HTTP 429). Превышен лимит обращений — попробуйте позже или уменьшите объём.",
        451: "Геоблокировка (HTTP 451). API доступен только с российских IP-адресов.",
        500: "Ошибка на стороне сервера (HTTP 500). Попробуйте позже.",
        502: "Сервер временно недоступен (HTTP 502). Попробуйте позже.",
        503: "Сервер перегружен (HTTP 503). Попробуйте позже.",
    }
    if status in messages:
        detail = ""
        if body:
            try:
                err_msg = json.loads(body).get("error_message", "")
                if err_msg:
                    detail = f" Сервер: «{err_msg}»"
            except Exception:
                if len(body) <= 200:
                    detail = f" {body}"
        return messages[status] + detail
    if 500 <= status < 600:
        return f"Ошибка сервера (HTTP {status}). Попробуйте позже."
    return f"Ошибка HTTP {status}."


# ══════════════════════════════════════════════════════════════════════
# Основное окно
# ══════════════════════════════════════════════════════════════════════

class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(600, 400)
        self.root.configure(bg=COLOR_BG)

        # Переменные
        self.codes: list[str] = []
        self.true_results: dict[str, dict] | None = None   # {code: raw_data}
        self.parsed_rows: list[list[str]] | None = None    # для Excel
        self.is_processing = False
        self._stop_requested = False                        # флаг остановки
        self.output_path: Path | None = None
        self._pending_update: dict | None = None            # инфо об обновлении (если есть)

        # Загружаем .env
        load_env(SCRIPT_DIR)

        # Подключаем логирование для модуля авторизации
        set_auth_log_fn(log_to_gui)

        # Строим интерфейс
        self._build_menu()
        self._build_layout()

        # Запускаем обработку очереди логов
        self._poll_log_queue()

        # Проверяем наличие токена
        token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        if not token:
            log_to_gui("⚠ Токен не задан. Открой «Настройки → Токен» и вставь токен.", "warn")

        # Автопроверка обновлений при старте
        self.root.after(1500, self._auto_check_updates)

    # ── Меню ────────────────────────────────────────────────────────

    def _build_menu(self) -> None:
        menubar = Menu(self.root, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                       activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)

        # Файл
        file_menu = Menu(menubar, tearoff=0,
                         bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                         activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)
        file_menu.add_command(label="📂 Загрузить коды из файла...", command=self._load_codes_from_file)
        file_menu.add_separator()
        file_menu.add_command(label="🚪 Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)

        # Настройки
        settings_menu = Menu(menubar, tearoff=0,
                             bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                             activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)
        settings_menu.add_command(label="🔑 Токен True API...", command=self._open_token_dialog)
        settings_menu.add_command(label="🔐 Получить токен через УКЭП...", command=self._open_ukep_dialog)
        menubar.add_cascade(label="Настройки", menu=settings_menu)

        # Справка
        help_menu = Menu(menubar, tearoff=0,
                         bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                         activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)
        help_menu.add_command(label="📖 Инструкция", command=self._open_help)
        help_menu.add_separator()
        help_menu.add_command(label="🔄 Проверить обновления...", command=self._check_updates)
        help_menu.add_command(label="ℹ О программе", command=self._show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

        self.root.config(menu=menubar)

    # ── Макет ───────────────────────────────────────────────────────

    def _build_layout(self) -> None:
        # Верхняя панель
        top_frame = Frame(self.root, bg=COLOR_HEADER_BG, height=50)
        top_frame.pack(fill="x", side="top")
        top_frame.pack_propagate(False)

        title_label = Label(top_frame, text="Проверка кодов маркировки «CISChecker»",
                            font=("Segoe UI", 14, "bold"), bg=COLOR_HEADER_BG, fg=COLOR_BUTTON_FG)
        title_label.pack(side="left", padx=16, pady=12)

        # Строка статуса токена
        token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        status_text = "🔑 True API: ✓" if token else "🔑 True API: не настроен"
        status_color = COLOR_LOG_SUCCESS if token else COLOR_LOG_WARN
        self.token_status_label = Label(top_frame, text=status_text,
                                        font=("Segoe UI", 9), bg=COLOR_HEADER_BG, fg=status_color)
        self.token_status_label.pack(side="right", padx=16, pady=12)

        # Обновляем статус токена с учётом срока действия
        if token:
            self._update_token_status(token)

        # Кнопка-индикатор обновлений (появляется после проверки)
        self.btn_update = Button(top_frame, text="🔄 Проверка...",
                                 font=("Segoe UI", 9),
                                 bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE,
                                 activebackground=COLOR_UPDATE_NONE, activeforeground=COLOR_UPDATE_FG_NONE,
                                 relief="flat", padx=10, pady=2,
                                 state=DISABLED,
                                 command=self._on_update_button)
        self.btn_update.pack(side="right", padx=(0, 8), pady=12)

        # Панель кнопок
        btn_frame = Frame(self.root, bg=COLOR_FRAME_BG, height=60)
        btn_frame.pack(fill="x", side="top", padx=8, pady=(8, 0))
        btn_frame.pack_propagate(False)

        self.btn_load = Button(btn_frame, text="📂 Отправить список кодов",
                               font=("Segoe UI", 10),
                               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
                               relief="flat", padx=16, pady=6,
                               command=self._load_and_run)
        self.btn_load.pack(side="left", padx=8, pady=12)

        self.btn_export = Button(btn_frame, text="📥 Выгрузка XLSX",
                                 font=("Segoe UI", 10),
                                 bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                                 activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
                                 relief="flat", padx=16, pady=6,
                                 state=DISABLED,
                                 command=self._export_xlsx)
        self.btn_export.pack(side="left", padx=8, pady=12)

        self.btn_stop = Button(btn_frame, text="⏹ Стоп",
                               font=("Segoe UI", 10),
                               bg="#e64553", fg="#ffffff",
                               activebackground="#d63543", activeforeground="#ffffff",
                               relief="flat", padx=16, pady=6,
                               state=DISABLED,
                               command=self._stop_processing)
        self.btn_stop.pack(side="left", padx=8, pady=12)

        # Выпадающий список товарных групп
        pg_frame = Frame(btn_frame, bg=COLOR_FRAME_BG)
        pg_frame.pack(side="left", padx=(16, 4), pady=12, fill="y")

        pg_label = Label(pg_frame, text="ТГ:", font=("Segoe UI", 9),
                         bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO)
        pg_label.pack(side="left", padx=(0, 4))

        self.pg_var = StringVar(value=PRODUCT_GROUPS_DEFAULT)
        self.pg_combo = ttk.Combobox(pg_frame, textvariable=self.pg_var,
                                      values=list(PRODUCT_GROUPS.keys()),
                                      state="readonly", width=32,
                                      font=("Segoe UI", 9))
        self.pg_combo.pack(side="left")

        # Индикатор прогресса
        self.progress_var = StringVar(value="Готов")
        self.progress_label = Label(btn_frame, textvariable=self.progress_var,
                                    font=("Segoe UI", 9), bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO)
        self.progress_label.pack(side="right", padx=16, pady=12)

        # Окно логов
        log_frame = Frame(self.root, bg=COLOR_BG)
        log_frame.pack(fill="both", expand=True, padx=8, pady=8)

        self.log_text = Text(
            log_frame,
            font=("Cascadia Code", 9),
            bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
            insertbackground=COLOR_LOG_FG,
            relief="flat",
            padx=8, pady=8,
            wrap="word",
            state=DISABLED,
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Настройка тегов для цветного текста
        self.log_text.tag_config("info", foreground=COLOR_LOG_INFO)
        self.log_text.tag_config("success", foreground=COLOR_LOG_SUCCESS)
        self.log_text.tag_config("error", foreground=COLOR_LOG_ERROR)
        self.log_text.tag_config("warn", foreground=COLOR_LOG_WARN)
        self.log_text.tag_config("bold", font=("Cascadia Code", 9, "bold"))

    # ── Очередь логов ───────────────────────────────────────────────

    def _poll_log_queue(self) -> None:
        """Периодически забирает сообщения из очереди и выводит в лог."""
        try:
            while True:
                msg, tag = log_queue.get_nowait()
                self._append_log(msg, tag)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _append_log(self, message: str, tag: str = "info") -> None:
        """Добавляет строку в окно логов."""
        self.log_text.config(state=NORMAL)
        if self.log_text.index("end-1c") != "1.0":
            self.log_text.insert(END, "\n")
        self.log_text.insert(END, message, tag)
        self.log_text.see(END)
        # Ограничиваем количество строк
        line_count = int(self.log_text.index("end-1c").split(".")[0])
        if line_count > LOG_MAX_LINES:
            self.log_text.delete("1.0", f"{line_count - LOG_MAX_LINES}.0")
        self.log_text.config(state=DISABLED)

    # ── Загрузка и запуск ──────────────────────────────────────────

    def _load_codes_from_file(self) -> None:
        """Диалог выбора файла с кодами."""
        path = filedialog.askopenfilename(
            title="Выберите файл с кодами маркировки",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if path:
            self._run_from_file(Path(path))

    def _load_and_run(self) -> None:
        """Загружает коды из файла и запускает проверку."""
        # Если уже есть коды в поле ввода (на будущее), используем их
        self._load_codes_from_file()

    def _run_from_file(self, file_path: Path) -> None:
        """Запускает проверку из файла."""
        if self.is_processing:
            messagebox.showwarning("Идёт обработка", "Дождитесь завершения текущей проверки.")
            return

        try:
            raw = file_path.read_text("utf-8")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")
            return

        codes = [line.strip() for line in raw.splitlines() if line.strip()]
        if not codes:
            messagebox.showwarning("Пустой файл", "Файл не содержит ни одного кода.")
            return

        # Дедупликация
        seen = set()
        unique = [c for c in codes if not (c in seen or seen.add(c))]
        self.codes = unique

        self._run_check()

    def _run_check(self) -> None:
        """Запускает процесс проверки в фоновом потоке."""
        token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        pg_name = self.pg_var.get()
        pg_code = PRODUCT_GROUPS.get(pg_name, "lp")
        if not token:
            messagebox.showwarning("Нет токена",
                                   "Токен True API не задан.\n\n"
                                   "Будет использован публичный API (без данных о владельце и без привязки к товарной группе).\n"
                                   "Для полной проверки с учётом ТГ откройте «Настройки → Токен».")
            self._start_public_thread()
        else:
            self._start_true_thread()

    def _quick_auth_check(self, token: str, pg_code: str = "lp") -> str | None:
        """
        Быстрая проверка токена одним запросом.
        Возвращает None если токен валиден, или текст ошибки.
        """
        test_codes = ["0102901036818042215U)lMHIaW2qGO"]  # любой код, даже невалидный
        from check_codes import http_post, TRUE_API
        url = f"{TRUE_API}?pg={pg_code}"
        status, body = http_post(
            url, json.dumps(test_codes),
            {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Authorization": f"Bearer {token}",
            },
            debug=False,
        )
        if status == 401:
            try:
                err = json.loads(body).get("error_message", body)
            except Exception:
                err = body or "неизвестная ошибка"
            return f"Токен недействителен (HTTP 401). Сервер ответил: {err}"
        if status == 403:
            return "Доступ запрещён (HTTP 403). Проверьте права токена."
        if status == 451:
            return "Геоблокировка (HTTP 451). API доступен только с российских IP."
        if status == 429:
            return "Слишком много запросов (HTTP 429). Подождите и повторите."
        if status is None or body is None:
            return "Не удалось подключиться к API. Проверьте интернет и российский IP."
        # 200 или 404 — ок, токен рабочий (404 — просто код не нашёлся)
        return None

    def _start_true_thread(self) -> None:
        """Запускает True API в фоновом потоке батчами по BATCH_SIZE с предварительной проверкой токена."""
        self._stop_requested = False
        self._set_processing_state(True)

        # Определяем выбранную товарную группу
        pg_name = self.pg_var.get()
        pg_code = PRODUCT_GROUPS.get(pg_name, "lp")

        log_to_gui("=" * 60, "bold")
        log_to_gui(f"Запуск проверки (True API). Кодов: {len(self.codes)}", "info")
        log_to_gui(f"Товарная группа: {pg_name} (pg={pg_code})", "info")
        log_to_gui(f"Размер батча: {BATCH_SIZE} кодов", "info")
        log_to_gui("=" * 60, "bold")

        token = os.environ.get("CHESTNYZNAK_TOKEN", "")

        def worker():
            # Шаг 0: быстрая проверка токена
            log_to_gui("⏳ Проверка токена...", "info")
            auth_error = self._quick_auth_check(token, pg_code)
            if auth_error:
                log_to_gui(f"❌ {auth_error}", "error")
                log_to_gui("=" * 60, "bold")
                log_to_gui("ПРОВЕРКА ПРЕРВАНА: неверный или просроченный токен.", "bold")
                log_to_gui("Обновите токен через «Настройки → Токен» и повторите.", "info")
                log_to_gui("=" * 60, "bold")
                self.true_results = None
                self.parsed_rows = None
                self.root.after(0, self._finish_processing)
                return

            log_to_gui("✓ Токен валиден. Начинаю проверку кодов...", "success")

            try:
                # Результаты: {code: item_dict}
                results: dict[str, dict] = {}
                total = len(self.codes)
                checked = 0

                # Разбиваем на батчи
                for batch_start in range(0, total, BATCH_SIZE):
                    if self._stop_requested:
                        log_to_gui("⏹ Остановлено по требованию пользователя.", "warn")
                        break

                    batch = self.codes[batch_start:batch_start + BATCH_SIZE]
                    batch_num = batch_start // BATCH_SIZE + 1
                    total_batches = (total + BATCH_SIZE - 1) // BATCH_SIZE
                    log_to_gui(f"  📦 Батч {batch_num}/{total_batches} ({len(batch)} кодов)...", "info")

                    status, data = true_check_with_retry_pg(
                        batch, pg_code, token, log_fn=log_to_gui,
                    )

                    if data is not None:
                        found = set()
                        for item in data:
                            cis_info = item.get("cisInfo", item)
                            c = cis_info.get("requestedCis", cis_info.get("cis", ""))
                            if c:
                                found.add(c)
                            results[c] = item
                        # Коды из батча, которые не вернулись
                        unmatched = 0
                        for code in batch:
                            if code not in found:
                                results[code] = results.get(code, {})
                                unmatched += 1
                            checked += 1
                        log_to_gui(f"  ✓ Батч {batch_num}/{total_batches}: получено {len(data)} результатов", "success")
                        # Предупреждение о возможной неверной ТГ
                        if found and unmatched == len(batch):
                            log_to_gui(f"  ⚠ Ни один код из батча не распознан. Возможно выбрана неверная товарная группа!", "warn")
                        self.root.after(0, self.progress_var.set, f"Проверка: {checked}/{total}")
                    else:
                        err_msg = explain_api_error(status, None)
                        for code in batch:
                            results[code] = {"error": f"True API: {err_msg}"}
                        checked += len(batch)
                        log_to_gui(f"  ✗ Батч {batch_num}/{total_batches}: {err_msg}", "error")
                        # Если ошибка авторизации — нет смысла продолжать
                        if status in (401, 403, 429, 451):
                            log_to_gui("  ⏹ Дальнейшие запросы бессмыслены — прерываю.", "error")
                            # Заполняем оставшиеся коды ошибкой
                            remaining = self.codes[batch_start + BATCH_SIZE:]
                            for code in remaining:
                                results[code] = {"error": f"True API: {err_msg}"}
                            break
                        self.root.after(0, self.progress_var.set, f"Проверка: {checked}/{total}")

                    time.sleep(0.15)

                self.true_results = results
                # Парсим строки для Excel
                self.parsed_rows = []
                for code in self.codes:
                    item = results.get(code, {"error": "Нет данных"})
                    row = parse_result(code, item, "true")
                    self.parsed_rows.append(row)
                    # Лог
                    if "error" in item:
                        log_to_gui(f"  ✗ {code[:40]}... → ОШИБКА: {item['error']}", "error")
                    else:
                        cis_info = item.get("cisInfo", item)
                        owner = cis_info.get("ownerName", "")
                        status_val = cis_info.get("status", "?")
                        product = cis_info.get("productName", "")
                        icon = "✓" if owner else "○"
                        log_to_gui(f"  {icon} {code[:40]}... → {status_val} | {product[:40]}", "success" if owner else "info")

            except Exception as e:
                log_to_gui(f"  ❌ Критическая ошибка: {e}", "error")
                self.true_results = None
                self.parsed_rows = None

            if self._stop_requested:
                log_to_gui("⏹ Проверка остановлена пользователем.", "warn")

            lines_ok = sum(1 for r in (self.parsed_rows or []) if not r[4].startswith("ОШИБКА"))
            lines_err = sum(1 for r in (self.parsed_rows or []) if r[4].startswith("ОШИБКА"))
            with_owner = len({r[6].strip() for r in (self.parsed_rows or []) if r[6] and r[6].strip()})
            log_to_gui("=" * 60, "bold")
            log_to_gui(f"ГОТОВО. Проверено: {len(self.codes)} | Успешно: {lines_ok} | Уникальных владельцев: {with_owner} | Ошибок: {lines_err}", "bold")
            log_to_gui("=" * 60, "bold")

            self.root.after(0, self._finish_processing)

        threading.Thread(target=worker, daemon=True).start()

    def _start_public_thread(self) -> None:
        """Запускает публичный API в фоновом потоке с прогрессом и остановкой."""
        self._stop_requested = False
        self._set_processing_state(True)
        log_to_gui("=" * 60, "bold")
        log_to_gui(f"Запуск проверки (Публичный API). Кодов: {len(self.codes)}", "info")
        log_to_gui("=" * 60, "bold")

        def worker():
            self.parsed_rows = []
            total = len(self.codes)
            for i, code in enumerate(self.codes, 1):
                if self._stop_requested:
                    log_to_gui(f"⏹ Остановлено на коде {i}/{total}", "warn")
                    break
                self.root.after(0, self.progress_var.set, f"Проверка: {i}/{total}")
                data = public_check(code)
                if data:
                    row = parse_result(code, data, "public")
                    self.parsed_rows.append(row)
                    status = data.get("outerStatus") or data.get("status", "?")
                    product = data.get("productName", "")
                    log_to_gui(f"  [{i}/{total}] ✓ {code[:35]}... → {status} | {product[:40]}", "info")
                else:
                    # Проверяем последнюю ошибку из http_post — пытаемся дать пояснение
                    self.parsed_rows.append([code, "", "", "", "ОШИБКА: нет ответа от сервера", "", "", "", "", ""])
                    log_to_gui(f"  [{i}/{total}] ✗ {code[:35]}... → нет ответа от сервера", "error")
                time.sleep(0.15)

            if self._stop_requested:
                log_to_gui("⏹ Проверка остановлена пользователем.", "warn")

            lines_ok = sum(1 for r in self.parsed_rows if not r[4].startswith("ОШИБКА"))
            lines_err = sum(1 for r in self.parsed_rows if r[4].startswith("ОШИБКА"))
            log_to_gui("=" * 60, "bold")
            log_to_gui(f"ГОТОВО. Проверено: {len(self.parsed_rows)}/{total} | Успешно: {lines_ok} | Ошибок: {lines_err}", "bold")
            log_to_gui("=" * 60, "bold")

            self.root.after(0, self._finish_processing)

        threading.Thread(target=worker, daemon=True).start()

    def _set_processing_state(self, processing: bool) -> None:
        """Переключает интерфейс между состояниями «идёт обработка» и «готов». """
        self.is_processing = processing
        state = DISABLED if processing else NORMAL
        self.btn_load.config(state=state)
        self.btn_stop.config(state=NORMAL if processing else DISABLED)

        if processing:
            self.btn_export.config(state=DISABLED)
            self.progress_var.set("Выполняется...")
        else:
            self.progress_var.set("Готов")

    def _finish_processing(self) -> None:
        """Вызывается из основного потока после завершения фоновой работы."""
        self._set_processing_state(False)
        if self.parsed_rows:
            self.btn_export.config(state=NORMAL)
            lines_ok = sum(1 for r in self.parsed_rows if not r[4].startswith("ОШИБКА"))
            self.progress_var.set(f"Готово: {lines_ok}/{len(self.codes)}")

    def _stop_processing(self) -> None:
        """Пользователь нажал Стоп. Устанавливаем флаг для потока."""
        self._stop_requested = True
        log_to_gui("⏹ Запрошена остановка...", "warn")

    # ── Экспорт в Excel ─────────────────────────────────────────────

    def _export_xlsx(self) -> None:
        """Сохраняет результаты в Excel."""
        if not self.parsed_rows:
            messagebox.showwarning("Нет данных", "Сначала выполните проверку кодов.")
            return

        path = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".xlsx",
            initialfile="result.xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("CSV файлы", "*.csv")]
        )
        if not path:
            return

        try:
            save_excel(self.parsed_rows, path)
            log_to_gui(f"✓ Результат сохранён: {path}", "success")
            messagebox.showinfo("Готово", f"Файл сохранён:\n{path}")
        except Exception as e:
            log_to_gui(f"❌ Ошибка сохранения: {e}", "error")
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    # ── Диалоги ─────────────────────────────────────────────────────

    def _open_token_dialog(self) -> None:
        """Окно настроек токена."""
        dlg = Toplevel(self.root)
        dlg.title("Настройки — Токен True API")
        dlg.geometry("520x260")
        dlg.configure(bg=COLOR_FRAME_BG)
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        Label(dlg, text="Токен True API",
              font=("Segoe UI", 12, "bold"), bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(pady=(16, 4))

        Label(dlg,
              text="Войдите в ЛК markirovka.crpt.ru → F12 → Console → localStorage.getItem('token')",
              font=("Segoe UI", 9), bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO, wraplength=480).pack(pady=(0, 8))

        token_entry = Entry(dlg, font=("Cascadia Code", 10),
                            bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                            insertbackground=COLOR_LOG_FG,
                            relief="flat", width=60)
        token_entry.pack(padx=16, pady=(0, 8), fill="x")

        current_token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        if current_token:
            token_entry.insert(0, current_token)

        # Фокус на поле ввода
        token_entry.focus_set()
        token_entry.icursor(len(current_token))  # курсор в конец

        def save_token() -> None:
            new_token = token_entry.get().strip()
            if not new_token:
                messagebox.showwarning("Пустой токен", "Введите токен или нажмите Отмена.", parent=dlg)
                return

            # Сохраняем в .env
            env_path = SCRIPT_DIR / ".env"
            lines = []
            found = False
            if env_path.exists():
                for line in env_path.read_text("utf-8").splitlines():
                    if line.startswith("CHESTNYZNAK_TOKEN="):
                        lines.append(f"CHESTNYZNAK_TOKEN={new_token}")
                        found = True
                    else:
                        lines.append(line)
            if not found:
                lines.append(f"CHESTNYZNAK_TOKEN={new_token}")
            env_path.write_text("\n".join(lines) + "\n", "utf-8")

            os.environ["CHESTNYZNAK_TOKEN"] = new_token
            # Сохраняем срок действия токена
            expires_at = self._get_token_expiry(new_token)
            exp_path = SCRIPT_DIR / ".token_expires"
            if expires_at is not None:
                exp_path.write_text(str(expires_at), "utf-8")
            else:
                # Для UUID-токенов: оценка ~10 часов по документации
                exp_path.write_text(str(time.time() + 36000), "utf-8")
            self._update_token_status(new_token)
            log_to_gui("🔑 Токен сохранён в .env", "success")
            dlg.destroy()

        btn_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        btn_frame.pack(pady=12)

        Button(btn_frame, text="💾 Сохранить",
               font=("Segoe UI", 10),
               bg=COLOR_BUTTON_ACCENT, fg="#1e1e2e",
               activebackground=COLOR_BUTTON_ACCENT_HOVER, activeforeground="#1e1e2e",
               relief="flat", padx=20, pady=6,
               command=save_token).pack(side="left", padx=8)

        Button(btn_frame, text="🗑 Очистить",
               font=("Segoe UI", 10),
               bg="#e64553", fg="#ffffff",
               activebackground="#d63543", activeforeground="#ffffff",
               relief="flat", padx=20, pady=6,
               command=lambda: token_entry.delete(0, END)).pack(side="left", padx=8)

        Button(btn_frame, text="Отмена",
               font=("Segoe UI", 10),
               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=20, pady=6,
               command=dlg.destroy).pack(side="left", padx=8)

    def _update_token_status(self, token: str) -> None:
        """Обновляет индикатор статуса токена с проверкой срока действия."""
        if not token:
            self.token_status_label.config(text="🔑 True API: не настроен", fg=COLOR_LOG_WARN)
            return

        # Определяем срок действия токена
        expires_at = self._get_token_expiry(token)
        now = time.time()

        if expires_at is not None:
            remaining_h = (expires_at - now) / 3600
            remaining_min = (expires_at - now) / 60
            if remaining_min <= 0:
                self.token_status_label.config(
                    text=f"🔑 Токен: ❌ просрочен",
                    fg=COLOR_LOG_ERROR
                )
                log_to_gui("❌ Токен просрочен. Получите новый через «Настройки → Получить токен через УКЭП».", "error")
            elif remaining_h < 1:
                self.token_status_label.config(
                    text=f"🔑 Токен: ⚠ {remaining_min:.0f} мин",
                    fg=COLOR_LOG_WARN
                )
                log_to_gui(f"⚠ Токен истекает через {remaining_min:.0f} мин. Обновите его.", "warn")
            else:
                exp_str = datetime.fromtimestamp(expires_at, tz=timezone(timedelta(hours=3))).strftime("%H:%M")
                self.token_status_label.config(
                    text=f"🔑 Токен: ✓ до {exp_str} МСК",
                    fg=COLOR_LOG_SUCCESS
                )
        else:
            # Не удалось определить срок из токена — читаем из файла
            exp_path = SCRIPT_DIR / ".token_expires"
            if exp_path.exists():
                try:
                    expires_ts = float(exp_path.read_text("utf-8").strip())
                    remaining_h = (expires_ts - now) / 3600
                    remaining_min = (expires_ts - now) / 60
                    if remaining_min <= 0:
                        self.token_status_label.config(
                            text=f"🔑 Токен: ❌ просрочен",
                            fg=COLOR_LOG_ERROR
                        )
                        log_to_gui("❌ Токен просрочен. Получите новый через «Настройки → Получить токен через УКЭП».", "error")
                    elif remaining_h < 1:
                        self.token_status_label.config(
                            text=f"🔑 Токен: ⚠ {remaining_min:.0f} мин",
                            fg=COLOR_LOG_WARN
                        )
                        log_to_gui(f"⚠ Токен истекает через {remaining_min:.0f} мин.", "warn")
                    else:
                        exp_str = datetime.fromtimestamp(expires_ts, tz=timezone(timedelta(hours=3))).strftime("%H:%M")
                        self.token_status_label.config(
                            text=f"🔑 Токен: ✓ до {exp_str} МСК",
                            fg=COLOR_LOG_SUCCESS
                        )
                except (ValueError, OSError):
                    self.token_status_label.config(text="🔑 True API: ✓", fg=COLOR_LOG_SUCCESS)
            else:
                # Нет информации о сроке — просто показываем ✓
                self.token_status_label.config(text="🔑 True API: ✓", fg=COLOR_LOG_SUCCESS)

    @staticmethod
    def _get_token_expiry(token: str) -> float | None:
        """
        Извлекает время истечения токена.
        - JWT: декодирует payload, возвращает exp (unix timestamp)
        - UUID-like: возвращает None (срок неизвестен)
        """
        try:
            # JWT формат: header.payload.signature
            parts = token.split(".")
            if len(parts) == 3:
                # Декодируем payload (добавляем padding если нужно)
                payload_b64 = parts[1]
                # Base64url → standard base64
                payload_b64 += "=" * (4 - len(payload_b64) % 4)
                import base64
                payload_json = base64.urlsafe_b64decode(payload_b64)
                payload = json.loads(payload_json)
                exp = payload.get("exp")
                if exp is not None:
                    return float(exp)
        except Exception:
            pass

        # UUID-подобный токен — срок неизвестен
        return None

    def _open_ukep_dialog(self) -> None:
        """Диалог получения токена через УКЭП (электронная подпись)."""
        dlg = Toplevel(self.root)
        dlg.title("Получить токен через УКЭП")
        dlg.geometry("600x500")
        dlg.configure(bg=COLOR_FRAME_BG)
        dlg.resizable(True, True)
        dlg.transient(self.root)
        dlg.grab_set()

        # Заголовок
        Label(dlg, text="🔐 Авторизация через УКЭП",
              font=("Segoe UI", 13, "bold"), bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(pady=(16, 4))
        Label(dlg,
              text="Подключите USB-токен (RuToken) и выберите сертификат.\n"
                   "Токен будет получен автоматически и сохранён в настройки.",
              font=("Segoe UI", 9), bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO,
              wraplength=560, justify="center").pack(pady=(0, 12))

        # Список сертификатов
        certs_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        certs_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        Label(certs_frame, text="Сертификаты УКЭП:", font=("Segoe UI", 10),
              bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(anchor="w", pady=(0, 4))

        certs_listbox_frame = Frame(certs_frame, bg=COLOR_LOG_BG)
        certs_listbox_frame.pack(fill="both", expand=True)

        certs_listbox = Listbox(certs_listbox_frame, font=("Cascadia Code", 9),
                                 bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                                 selectbackground=COLOR_BUTTON_ACCENT,
                                 selectforeground="#1e1e2e",
                                 relief="flat", height=6)
        certs_scrollbar = Scrollbar(certs_listbox_frame, orient=VERTICAL,
                                     command=certs_listbox.yview)
        certs_listbox.config(yscrollcommand=certs_scrollbar.set)
        certs_listbox.pack(side="left", fill="both", expand=True)
        certs_scrollbar.pack(side="right", fill="y")

        # Загрузка сертификатов
        certs_data: list[dict] = []

        def load_certs() -> None:
            """Загружает список сертификатов в Listbox."""
            certs_listbox.delete(0, END)
            certs_data.clear()
            try:
                certs = list_certificates()
                if not certs:
                    certs_listbox.insert(END, "  Сертификаты не найдены.")
                    certs_listbox.insert(END, "  Убедитесь, что КриптоПро CSP установлен")
                    certs_listbox.insert(END, "  и УКЭП (RuToken) подключён.")
                else:
                    for cert in certs:
                        subject = cert.get("subject", "Неизвестный")
                        inn = cert.get("inn", "")
                        thumbprint = cert.get("thumbprint", "")
                        display = f"  {subject}"
                        if inn:
                            display += f" (ИНН: {inn})"
                        certs_listbox.insert(END, display)
                        certs_data.append(cert)
            except Exception as e:
                certs_listbox.insert(END, f"  Ошибка: {e}")

        # Кнопка обновления сертификатов
        refresh_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        refresh_frame.pack(fill="x", padx=16, pady=(0, 4))

        Button(refresh_frame, text="🔄 Обновить список",
               font=("Segoe UI", 9), bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=12, pady=4,
               command=load_certs).pack(side="left")

        # Статус
        status_var = StringVar(value="")

        status_label = Label(refresh_frame, textvariable=status_var,
                             font=("Segoe UI", 9), bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO)
        status_label.pack(side="right", padx=4)

        # Метод авторизации: JWT через УКЭП
        method_label_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        method_label_frame.pack(fill="x", padx=16, pady=(0, 8))

        Label(method_label_frame, text="Метод: JWT через УКЭП",
              font=("Segoe UI", 10), bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO).pack(anchor="w")

        # Кнопки действия
        btn_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        btn_frame.pack(pady=(0, 16))

        def do_auth() -> None:
            """Запускает авторизацию через УКЭП в фоновом потоке."""
            # Определяем thumbprint выбранного сертификата
            thumbprint = ""
            sel = certs_listbox.curselection()
            if sel and sel[0] < len(certs_data):
                thumbprint = certs_data[sel[0]].get("thumbprint", "")

            status_var.set("⏳ Авторизация...")
            dlg_btn_auth.config(state=DISABLED)

            def worker() -> None:
                success, result = auth_jwt(thumbprint)

                def on_done() -> None:
                    dlg_btn_auth.config(state=NORMAL)
                    if success:
                        # Сохраняем токен
                        new_token = result
                        env_path = SCRIPT_DIR / ".env"
                        lines = []
                        found = False
                        if env_path.exists():
                            for line in env_path.read_text("utf-8").splitlines():
                                if line.startswith("CHESTNYZNAK_TOKEN="):
                                    lines.append(f"CHESTNYZNAK_TOKEN={new_token}")
                                    found = True
                                else:
                                    lines.append(line)
                        if not found:
                            lines.append(f"CHESTNYZNAK_TOKEN={new_token}")
                        # Сохраняем ИНН из сертификата (для автоподстановки)
                        inn_from_cert = ""
                        if sel and sel[0] < len(certs_data):
                            inn_from_cert = certs_data[sel[0]].get("inn", "")
                        inn_lines = [l for l in lines if not l.startswith("CHESTNYZNAK_INN=")]
                        if inn_from_cert:
                            inn_lines.append(f"CHESTNYZNAK_INN={inn_from_cert}")
                        env_path.write_text("\n".join(inn_lines) + "\n", "utf-8")

                        os.environ["CHESTNYZNAK_TOKEN"] = new_token
                        if inn_from_cert:
                            os.environ["CHESTNYZNAK_INN"] = inn_from_cert
                        # Сохраняем срок действия токена
                        expires_at = self._get_token_expiry(new_token)
                        exp_path = SCRIPT_DIR / ".token_expires"
                        if expires_at is not None:
                            exp_path.write_text(str(expires_at), "utf-8")
                        else:
                            exp_path.write_text(str(time.time() + 36000), "utf-8")
                        self._update_token_status(new_token)

                        token_preview = new_token[:20] + "..." if len(new_token) > 20 else new_token
                        status_var.set(f"✅ Токен получен: {token_preview}")
                        log_to_gui("🔐 JWT-токен получен через УКЭП", "success")
                        messagebox.showinfo("Токен получен",
                                            f"JWT-токен успешно получен и сохранён!",
                                            parent=dlg)
                        dlg.destroy()
                    else:
                        status_var.set("❌ Ошибка")
                        messagebox.showerror("Ошибка авторизации", result, parent=dlg)

                self.root.after(0, on_done)

            threading.Thread(target=worker, daemon=True).start()

        dlg_btn_auth = Button(btn_frame, text="🔐 Получить токен",
                               font=("Segoe UI", 11, "bold"),
                               bg=COLOR_BUTTON_ACCENT, fg="#1e1e2e",
                               activebackground=COLOR_BUTTON_ACCENT_HOVER, activeforeground="#1e1e2e",
                               relief="flat", padx=20, pady=8,
                               command=do_auth)
        dlg_btn_auth.pack(side="left", padx=8)

        Button(btn_frame, text="Отмена",
               font=("Segoe UI", 10),
               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=20, pady=8,
               command=dlg.destroy).pack(side="left", padx=8)

        # Загружаем сертификаты при открытии
        load_certs()

    def _open_help(self) -> None:
        """Показывает инструкцию."""
        dlg = Toplevel(self.root)
        dlg.title("Инструкция")
        dlg.geometry("600x450")
        dlg.configure(bg=COLOR_FRAME_BG)
        dlg.resizable(True, True)
        dlg.transient(self.root)

        text = Text(dlg, font=("Segoe UI", 10),
                    bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                    relief="flat", padx=12, pady=12, wrap="word")
        text.pack(fill="both", expand=True)

        instructions = """
        ЧЕСТНЫЙ ЗНАК — ПРОВЕРКА КОДОВ МАРКИРОВКИ

        Как пользоваться:

        1. Настройте токен True API (Настройки → Токен)
           - Войдите в ЛК: https://markirovka.crpt.ru
           - F12 → Console → localStorage.getItem('token')
           - Скопируйте и вставьте в окно настроек

        2. Загрузите список кодов (Файл → Загрузить коды)
           - TXT-файл, по одному коду на строку

        3. Нажмите «Отправить список кодов»
           - Скрипт автоматически определит товарную группу
           - И найдёт владельца для каждого кода

        4. Выгрузите результат в Excel / CSV

        Формат выгрузки:
        - Штрихкод
        - Индекс картинки
        - Статус
        - Количество
        - Владелец
        - Производитель
        - Дата ввода в оборот
        - Способ ввода в оборот

        Требования:
        - Российский IP (сервера ЧЗ геоблокируют зарубежные)
        - Python 3.8+
        - openpyxl (для Excel)
        """

        text.insert("1.0", instructions.strip())
        text.config(state=DISABLED)

    def _check_updates(self) -> None:
        """Проверка обновлений через GitHub Releases (в фоне)."""
        if self.is_processing:
            messagebox.showwarning("Обновление", "Сначала дождитесь завершения текущей проверки.")
            return
        self._run_update_check()

    def _auto_check_updates(self) -> None:
        """Автопроверка обновлений при старте (тихая, без диалогов)."""
        log_to_gui("🔄 Проверка обновлений...", "info")
        self.btn_update.config(text="🔄 Проверка...", state=DISABLED,
                                bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE)

        def _worker():
            has_update, message, release_info = _check_update()

            def _on_result():
                self._pending_update = release_info if has_update else None
                if has_update:
                    latest_ver = release_info.get("version", "?") if release_info else "?"
                    log_to_gui(f"🟢 Доступно обновление: v{latest_ver} (у вас v{APP_VERSION})", "success")
                    self.btn_update.config(
                        text=f"⬆ Обновление v{latest_ver}",
                        state=NORMAL,
                        bg=COLOR_UPDATE_AVAIL, fg=COLOR_UPDATE_FG_AVAIL,
                        activebackground="#8be89c", activeforeground=COLOR_UPDATE_FG_AVAIL,
                    )
                else:
                    log_to_gui("✓ Установлена последняя версия.", "info")
                    self.btn_update.config(
                        text="✓ Обновлений нет",
                        state=DISABLED,
                        bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE,
                        activebackground=COLOR_UPDATE_NONE, activeforeground=COLOR_UPDATE_FG_NONE,
                    )

            self.root.after(0, _on_result)

        threading.Thread(target=_worker, daemon=True).start()

    def _run_update_check(self) -> None:
        """Ручная проверка обновлений из меню (с диалогами)."""
        log_to_gui("🔄 Проверяю обновления...", "info")
        self.btn_update.config(text="🔄 Проверка...", state=DISABLED,
                                bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE)

        def _worker():
            has_update, message, release_info = _check_update()

            def _on_result():
                self._pending_update = release_info if has_update else None
                if has_update and release_info:
                    latest_ver = release_info.get("version", "?")
                    log_to_gui(f"🟢 Доступно обновление: v{latest_ver} (у вас v{APP_VERSION})", "success")
                    self.btn_update.config(
                        text=f"⬆ Обновление v{latest_ver}",
                        state=NORMAL,
                        bg=COLOR_UPDATE_AVAIL, fg=COLOR_UPDATE_FG_AVAIL,
                        activebackground="#8be89c", activeforeground=COLOR_UPDATE_FG_AVAIL,
                    )
                    result = messagebox.askyesno(
                        "Доступно обновление",
                        f"Доступна новая версия v{latest_ver} (у вас v{APP_VERSION}).\n\n"
                        f"Скачать и установить обновление?\n\n"
                        f"{'⚠ Приложение перезапустится.' if _is_frozen() else 'ℹ Автообновление доступно только в .exe-режиме.'}",
                    )
                    if result:
                        self._apply_update(release_info)
                else:
                    log_to_gui("✓ Установлена последняя версия.", "info")
                    self.btn_update.config(
                        text="✓ Обновлений нет",
                        state=DISABLED,
                        bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE,
                        activebackground=COLOR_UPDATE_NONE, activeforeground=COLOR_UPDATE_FG_NONE,
                    )
                    messagebox.showinfo("Обновление", message)

            self.root.after(0, _on_result)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_update_button(self) -> None:
        """Нажатие на кнопку-индикатор обновлений."""
        if self._pending_update:
            release_info = self._pending_update
            latest_ver = release_info.get("version", "?")
            result = messagebox.askyesno(
                "Обновление",
                f"Скачать и установить v{latest_ver}?\n\n"
                f"{'⚠ Приложение перезапустится.' if _is_frozen() else 'ℹ Автообновление доступно только в .exe-режиме.'}",
            )
            if result:
                self._apply_update(release_info)

    def _apply_update(self, release_info: dict) -> None:
        """Скачивает и применяет обновление."""
        exe_url = release_info.get("exe_url")
        if not exe_url:
            # Нет .exe в релизе — открываем браузер на страницу релиза
            import webbrowser
            tag = release_info.get("tag", "")
            url = f"https://github.com/{GITHUB_REPO}/releases/tag/{tag}"
            messagebox.showinfo(
                "Обновление",
                f".exe не найден в релизе.\n\n"
                f"Скачайте обновление вручную:\n{url}",
            )
            webbrowser.open(url)
            return

        if not _is_frozen():
            # В режиме .py — отправляем на GitHub
            import webbrowser
            tag = release_info.get("tag", "")
            url = f"https://github.com/{GITHUB_REPO}/releases/tag/{tag}"
            messagebox.showinfo(
                "Обновление",
                f"Автообновление работает только в .exe.\n\n"
                f"Скачайте исходники с GitHub:\n{url}",
            )
            webbrowser.open(url)
            return

        log_to_gui("📥 Скачиваю обновление...", "info")

        def _worker():
            success, msg = _perform_update(
                exe_url,
                progress_fn=lambda d, t: None,  # Прогресс скачивания (пока без GUI)
            )

            def _on_result():
                if success:
                    log_to_gui("✅ Обновление скачано. Перезапускаю...", "bold")
                    # Закрываем приложение — .bat скрипт подменит .exe и запустит заново
                    self.root.after(1000, self.root.destroy)
                else:
                    log_to_gui(f"❌ {msg}", "error")
                    messagebox.showerror("Ошибка обновления", msg)

            self.root.after(0, _on_result)

        threading.Thread(target=_worker, daemon=True).start()

    def _show_about(self) -> None:
        """Диалог «О программе»."""
        messagebox.showinfo(
            "О программе",
            f"{APP_TITLE}\n\n"
            f"Версия: {APP_VERSION}\n"
            f"Python + Tkinter\n\n"
            f"Использует публичный и True API\n"
            f"системы маркировки «Честный Знак» (CISChecker).\n\n"
            f"Репозиторий: github.com/{GITHUB_REPO}\n"
            f"© 2026"
        )


# ══════════════════════════════════════════════════════════════════════
# Точка входа
# ══════════════════════════════════════════════════════════════════════

def main() -> None:
    root = Tk()
    app = App(root)
    log_to_gui(f"🚀 {APP_TITLE} v{APP_VERSION}", "bold")
    log_to_gui(f"📁 Папка проекта: {SCRIPT_DIR}", "info")
    log_to_gui("📋 Готов к работе. Загрузите файл с кодами или выберите «Файл → Загрузить коды».", "info")
    root.mainloop()


if __name__ == "__main__":
    main()
