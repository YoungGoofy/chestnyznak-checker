"""
Главное окно приложения CISChecker (tkinter).
"""
from __future__ import annotations

import json
import os
import threading
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from tkinter import (
    Tk, Frame, Label, Button, Menu,
    filedialog, messagebox, StringVar,
    DISABLED, NORMAL,
)
from tkinter import ttk

from .. import APP_TITLE, APP_VERSION, GITHUB_REPO
from ..core.constants import PRODUCT_GROUPS, PRODUCT_GROUPS_DEFAULT, BATCH_SIZE
from ..core.env import load_env
from ..core.checker import (
    true_check_with_retry_pg, public_check, explain_http_status,
)
from ..core.parser import parse_result
from ..core.excel import save_excel
from ..core.api import http_post
from ..core.constants import TRUE_API

from .theme import *
from .log_widget import LogWidget, log_to_gui
from .updater_ui import UpdaterUI
from .dialogs import (
    TokenDialog, UkepDialog, HelpDialog, show_about, get_token_expiry,
)


def get_app_dir() -> Path:
    """Папка приложения (работает и в .py, и в .exe)."""
    import sys
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent.parent


SCRIPT_DIR = get_app_dir()


class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(600, 400)
        self.root.configure(bg=COLOR_BG)

        # Состояние
        self.codes: list[str] = []
        self.parsed_rows: list[list[str]] | None = None
        self.is_processing = False
        self._stop_requested = False

        # Загружаем .env
        load_env(SCRIPT_DIR)

        # Строим интерфейс
        self._build_menu()
        self._build_layout()

        # Запуск опроса логов
        self.log_widget.poll_queue(root)

        # Проверяем токен
        token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        if not token:
            log_to_gui("⚠ Токен не задан. Откройте «Настройки → Токен» или «Получить через УКЭП».", "warn")

        # Автопроверка обновлений
        self.root.after(1500, self.updater_ui.auto_check)

    # ── Меню ────────────────────────────────────────────────────────

    def _build_menu(self) -> None:
        menubar = Menu(self.root, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                       activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)

        file_menu = Menu(menubar, tearoff=0, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                         activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)
        file_menu.add_command(label="📂 Загрузить коды из файла...", command=self._load_codes_from_file)
        file_menu.add_separator()
        file_menu.add_command(label="🚪 Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)

        settings_menu = Menu(menubar, tearoff=0, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                             activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)
        settings_menu.add_command(label="🔑 Токен True API...", command=self._open_token_dialog)
        settings_menu.add_command(label="🔐 Получить токен через УКЭП...", command=self._open_ukep_dialog)
        menubar.add_cascade(label="Настройки", menu=settings_menu)

        help_menu = Menu(menubar, tearoff=0, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG,
                         activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG)
        help_menu.add_command(label="📖 Инструкция", command=lambda: HelpDialog(self.root))
        help_menu.add_separator()
        help_menu.add_command(label="🔄 Проверить обновления...", command=self._check_updates)
        help_menu.add_command(label="ℹ О программе", command=lambda: show_about(self.root))
        menubar.add_cascade(label="Справка", menu=help_menu)

        self.root.config(menu=menubar)

    # ── Макет ───────────────────────────────────────────────────────

    def _build_layout(self) -> None:
        # Верхняя панель
        top_frame = Frame(self.root, bg=COLOR_HEADER_BG, height=50)
        top_frame.pack(fill="x", side="top")
        top_frame.pack_propagate(False)

        Label(top_frame, text="Проверка кодов маркировки «CISChecker»",
              font=FONT_TITLE, bg=COLOR_HEADER_BG, fg=COLOR_BUTTON_FG
              ).pack(side="left", padx=16, pady=12)

        # Статус токена
        token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        status_text = "🔑 True API: ✓" if token else "🔑 True API: не настроен"
        status_color = COLOR_LOG_SUCCESS if token else COLOR_LOG_WARN
        self.token_status_label = Label(top_frame, text=status_text,
                                         font=FONT_SMALL, bg=COLOR_HEADER_BG, fg=status_color)
        self.token_status_label.pack(side="right", padx=16, pady=12)

        if token:
            self._update_token_status(token)

        # Кнопка обновлений
        self.updater_ui = UpdaterUI(top_frame, self.root)

        # Панель кнопок
        btn_frame = Frame(self.root, bg=COLOR_FRAME_BG, height=60)
        btn_frame.pack(fill="x", side="top", padx=8, pady=(8, 0))
        btn_frame.pack_propagate(False)

        self.btn_load = Button(btn_frame, text="📂 Отправить список кодов",
                               font=FONT_BUTTON, bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
                               relief="flat", padx=16, pady=6,
                               command=self._load_and_run)
        self.btn_load.pack(side="left", padx=8, pady=12)

        self.btn_export = Button(btn_frame, text="📥 Выгрузка XLSX",
                                 font=FONT_BUTTON, bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
                                 activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
                                 relief="flat", padx=16, pady=6, state=DISABLED,
                                 command=self._export_xlsx)
        self.btn_export.pack(side="left", padx=8, pady=12)

        self.btn_stop = Button(btn_frame, text="⏹ Стоп",
                               font=FONT_BUTTON, bg="#e64553", fg="#ffffff",
                               activebackground="#d63543", activeforeground="#ffffff",
                               relief="flat", padx=16, pady=6, state=DISABLED,
                               command=self._stop_processing)
        self.btn_stop.pack(side="left", padx=8, pady=12)

        # Товарная группа
        pg_frame = Frame(btn_frame, bg=COLOR_FRAME_BG)
        pg_frame.pack(side="left", padx=(16, 4), pady=12, fill="y")

        Label(pg_frame, text="ТГ:", font=FONT_SMALL,
              bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO).pack(side="left", padx=(0, 4))

        self.pg_var = StringVar(value=PRODUCT_GROUPS_DEFAULT)
        self.pg_combo = ttk.Combobox(pg_frame, textvariable=self.pg_var,
                                      values=list(PRODUCT_GROUPS.keys()),
                                      state="readonly", width=32, font=FONT_SMALL)
        self.pg_combo.pack(side="left")

        # Прогресс
        self.progress_var = StringVar(value="Готов")
        Label(btn_frame, textvariable=self.progress_var,
              font=FONT_SMALL, bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO
              ).pack(side="right", padx=16, pady=12)

        # Логи
        log_frame = Frame(self.root, bg=COLOR_BG)
        log_frame.pack(fill="both", expand=True, padx=8, pady=8)
        self.log_widget = LogWidget(log_frame)
        self.log_widget.pack(fill="both", expand=True)

    # ── Загрузка и запуск ──────────────────────────────────────────

    def _load_codes_from_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите файл с кодами маркировки",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
        )
        if path:
            self._run_from_file(Path(path))

    def _load_and_run(self) -> None:
        self._load_codes_from_file()

    def _run_from_file(self, file_path: Path) -> None:
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

        seen = set()
        self.codes = [c for c in codes if not (c in seen or seen.add(c))]
        self._run_check()

    def _run_check(self) -> None:
        token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        if not token:
            messagebox.showwarning("Нет токена",
                                   "Токен True API не задан.\n\n"
                                   "Будет использован публичный API (без данных о владельце).\n"
                                   "Для полной проверки откройте «Настройки → Токен».")
            self._start_public_thread()
        else:
            self._start_true_thread()

    # ── True API ───────────────────────────────────────────────────

    def _quick_auth_check(self, token: str, pg_code: str = "lp") -> str | None:
        test_codes = ["0102901036818042215U)lMHIaW2qGO"]
        url = f"{TRUE_API}?pg={pg_code}"
        status, body = http_post(
            url, json.dumps(test_codes),
            {"Content-Type": "application/json", "Accept": "application/json",
             "Authorization": f"Bearer {token}"},
            debug=False,
        )
        if status == 401:
            try:
                err = json.loads(body).get("error_message", body)
            except Exception:
                err = body or "неизвестная ошибка"
            return f"Токен недействителен (HTTP 401). Сервер: {err}"
        if status == 403:
            return "Доступ запрещён (HTTP 403)."
        if status == 451:
            return "Геоблокировка (HTTP 451)."
        if status == 429:
            return "Слишком много запросов (HTTP 429)."
        if status is None or body is None:
            return "Не удалось подключиться к API."
        return None

    def _start_true_thread(self) -> None:
        self._stop_requested = False
        self._set_processing_state(True)

        pg_name = self.pg_var.get()
        pg_code = PRODUCT_GROUPS.get(pg_name, "lp")

        log_to_gui("=" * 60, "bold")
        log_to_gui(f"Запуск проверки (True API). Кодов: {len(self.codes)}", "info")
        log_to_gui(f"Товарная группа: {pg_name} (pg={pg_code})", "info")
        log_to_gui(f"Размер батча: {BATCH_SIZE}", "info")
        log_to_gui("=" * 60, "bold")

        token = os.environ.get("CHESTNYZNAK_TOKEN", "")

        def worker():
            log_to_gui("⏳ Проверка токена...", "info")
            auth_error = self._quick_auth_check(token, pg_code)
            if auth_error:
                log_to_gui(f"❌ {auth_error}", "error")
                log_to_gui("ПРОВЕРКА ПРЕРВАНА: неверный токен.", "bold")
                self.parsed_rows = None
                self.root.after(0, self._finish_processing)
                return

            log_to_gui("✓ Токен валиден.", "success")

            try:
                results: dict[str, dict] = {}
                total = len(self.codes)
                checked = 0

                for batch_start in range(0, total, BATCH_SIZE):
                    if self._stop_requested:
                        log_to_gui("⏹ Остановлено пользователем.", "warn")
                        break

                    batch = self.codes[batch_start:batch_start + BATCH_SIZE]
                    batch_num = batch_start // BATCH_SIZE + 1
                    total_batches = (total + BATCH_SIZE - 1) // BATCH_SIZE
                    log_to_gui(f"  📦 Батч {batch_num}/{total_batches} ({len(batch)} кодов)...", "info")

                    status, data = true_check_with_retry_pg(batch, pg_code, token, log_fn=log_to_gui)

                    if data is not None:
                        found = set()
                        for item in data:
                            cis_info = item.get("cisInfo", item)
                            c = cis_info.get("requestedCis", cis_info.get("cis", ""))
                            if c:
                                found.add(c)
                            results[c] = item
                        unmatched = 0
                        for code in batch:
                            if code not in found:
                                results[code] = results.get(code, {})
                                unmatched += 1
                            checked += 1
                        log_to_gui(f"  ✓ Батч {batch_num}/{total_batches}: {len(data)} результатов", "success")
                        if found and unmatched == len(batch):
                            log_to_gui("  ⚠ Ни один код не распознан. Возможно неверная ТГ!", "warn")
                        self.root.after(0, self.progress_var.set, f"Проверка: {checked}/{total}")
                    else:
                        err_msg = explain_http_status(status)
                        for code in batch:
                            results[code] = {"error": f"True API: {err_msg}"}
                        checked += len(batch)
                        log_to_gui(f"  ✗ Батч {batch_num}/{total_batches}: {err_msg}", "error")
                        if status in (401, 403, 429, 451):
                            log_to_gui("  ⏹ Прерываю — ошибка авторизации.", "error")
                            for code in self.codes[batch_start + BATCH_SIZE:]:
                                results[code] = {"error": f"True API: {err_msg}"}
                            break
                        self.root.after(0, self.progress_var.set, f"Проверка: {checked}/{total}")

                    time.sleep(0.05)  # уменьшено с 0.15

                # Парсим
                self.parsed_rows = []
                for code in self.codes:
                    item = results.get(code, {"error": "Нет данных"})
                    self.parsed_rows.append(parse_result(code, item, "true"))

            except Exception as e:
                log_to_gui(f"  ❌ Критическая ошибка: {e}", "error")
                self.parsed_rows = None

            lines_ok = sum(1 for r in (self.parsed_rows or []) if not r[4].startswith("ОШИБКА"))
            lines_err = sum(1 for r in (self.parsed_rows or []) if r[4].startswith("ОШИБКА"))
            log_to_gui("=" * 60, "bold")
            log_to_gui(f"ГОТОВО. Проверено: {len(self.codes)} | Успешно: {lines_ok} | Ошибок: {lines_err}", "bold")
            log_to_gui("=" * 60, "bold")
            self.root.after(0, self._finish_processing)

        threading.Thread(target=worker, daemon=True).start()

    def _start_public_thread(self) -> None:
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
                    self.parsed_rows.append(parse_result(code, data, "public"))
                    status = data.get("outerStatus") or data.get("status", "?")
                    log_to_gui(f"  [{i}/{total}] ✓ {code[:35]}... → {status}", "info")
                else:
                    self.parsed_rows.append([code, "", "", "", "ОШИБКА: нет ответа", "", "", "", "", ""])
                    log_to_gui(f"  [{i}/{total}] ✗ {code[:35]}... → нет ответа", "error")
                time.sleep(0.05)  # уменьшено с 0.15

            lines_ok = sum(1 for r in self.parsed_rows if not r[4].startswith("ОШИБКА"))
            lines_err = sum(1 for r in self.parsed_rows if r[4].startswith("ОШИБКА"))
            log_to_gui("=" * 60, "bold")
            log_to_gui(f"ГОТОВО. {len(self.parsed_rows)}/{total} | Успешно: {lines_ok} | Ошибок: {lines_err}", "bold")
            log_to_gui("=" * 60, "bold")
            self.root.after(0, self._finish_processing)

        threading.Thread(target=worker, daemon=True).start()

    # ── Управление состоянием ──────────────────────────────────────

    def _set_processing_state(self, processing: bool) -> None:
        self.is_processing = processing
        self.btn_load.config(state=DISABLED if processing else NORMAL)
        self.btn_stop.config(state=NORMAL if processing else DISABLED)
        if processing:
            self.btn_export.config(state=DISABLED)
            self.progress_var.set("Выполняется...")
        else:
            self.progress_var.set("Готов")

    def _finish_processing(self) -> None:
        self._set_processing_state(False)
        if self.parsed_rows:
            self.btn_export.config(state=NORMAL)
            lines_ok = sum(1 for r in self.parsed_rows if not r[4].startswith("ОШИБКА"))
            self.progress_var.set(f"Готово: {lines_ok}/{len(self.codes)}")

    def _stop_processing(self) -> None:
        self._stop_requested = True
        log_to_gui("⏹ Запрошена остановка...", "warn")

    # ── Экспорт ────────────────────────────────────────────────────

    def _export_xlsx(self) -> None:
        if not self.parsed_rows:
            messagebox.showwarning("Нет данных", "Сначала выполните проверку.")
            return
        path = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".xlsx", initialfile="result.xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("CSV файлы", "*.csv")],
        )
        if not path:
            return
        try:
            save_excel(self.parsed_rows, path)
            log_to_gui(f"✓ Сохранён: {path}", "success")
            messagebox.showinfo("Готово", f"Файл сохранён:\n{path}")
        except Exception as e:
            log_to_gui(f"❌ Ошибка: {e}", "error")
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    # ── Диалоги ────────────────────────────────────────────────────

    def _open_token_dialog(self) -> None:
        TokenDialog(self.root, SCRIPT_DIR, on_token_saved=self._update_token_status)

    def _open_ukep_dialog(self) -> None:
        UkepDialog(self.root, SCRIPT_DIR, on_token_saved=self._update_token_status)

    def _check_updates(self) -> None:
        if self.is_processing:
            messagebox.showwarning("Обновление", "Дождитесь завершения проверки.")
            return
        self.updater_ui.manual_check()

    def _update_token_status(self, token: str) -> None:
        if not token:
            self.token_status_label.config(text="🔑 True API: не настроен", fg=COLOR_LOG_WARN)
            return

        expires_at = get_token_expiry(token)
        now = time.time()

        if expires_at is not None:
            remaining_min = (expires_at - now) / 60
            remaining_h = remaining_min / 60
            if remaining_min <= 0:
                self.token_status_label.config(text="🔑 Токен: ❌ просрочен", fg=COLOR_LOG_ERROR)
                log_to_gui("❌ Токен просрочен. Обновите через УКЭП.", "error")
            elif remaining_h < 1:
                self.token_status_label.config(
                    text=f"🔑 Токен: ⚠ {remaining_min:.0f} мин", fg=COLOR_LOG_WARN)
            else:
                exp_str = datetime.fromtimestamp(
                    expires_at, tz=timezone(timedelta(hours=3))
                ).strftime("%H:%M")
                self.token_status_label.config(
                    text=f"🔑 Токен: ✓ до {exp_str} МСК", fg=COLOR_LOG_SUCCESS)
        else:
            # Проверяем файл .token_expires
            exp_path = SCRIPT_DIR / ".token_expires"
            if exp_path.exists():
                try:
                    expires_ts = float(exp_path.read_text("utf-8").strip())
                    remaining_min = (expires_ts - now) / 60
                    if remaining_min <= 0:
                        self.token_status_label.config(text="🔑 Токен: ❌ просрочен", fg=COLOR_LOG_ERROR)
                    elif remaining_min < 60:
                        self.token_status_label.config(
                            text=f"🔑 Токен: ⚠ {remaining_min:.0f} мин", fg=COLOR_LOG_WARN)
                    else:
                        exp_str = datetime.fromtimestamp(
                            expires_ts, tz=timezone(timedelta(hours=3))
                        ).strftime("%H:%M")
                        self.token_status_label.config(
                            text=f"🔑 Токен: ✓ до {exp_str} МСК", fg=COLOR_LOG_SUCCESS)
                except (ValueError, OSError):
                    self.token_status_label.config(text="🔑 True API: ✓", fg=COLOR_LOG_SUCCESS)
            else:
                self.token_status_label.config(text="🔑 True API: ✓", fg=COLOR_LOG_SUCCESS)


def main() -> None:
    root = Tk()
    app = App(root)
    log_to_gui(f"🚀 {APP_TITLE} v{APP_VERSION}", "bold")
    log_to_gui(f"📁 Папка: {SCRIPT_DIR}", "info")
    log_to_gui("📋 Готов. Загрузите файл с кодами.", "info")
    root.mainloop()
