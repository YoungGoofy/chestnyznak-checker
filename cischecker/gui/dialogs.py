"""
Диалоговые окна: Токен, УКЭП, Справка, О программе.
"""
from __future__ import annotations

import os
import threading
import time
from pathlib import Path
from tkinter import (
    Toplevel, Frame, Label, Button, Entry, Scrollbar,
    StringVar, END, VERTICAL, HORIZONTAL, DISABLED, NORMAL, messagebox,
)
from tkinter import ttk

from .. import APP_VERSION, APP_TITLE, GITHUB_REPO
from ..core.env import save_token_to_env
from ..auth.certificates import (
    list_all_valid_certificates, diagnose_com,
    set_log_fn as set_cert_log_fn,
)
from ..auth.jwt_flow import auth_jwt, set_log_fn as set_jwt_log_fn
from .theme import *
from .log_widget import log_to_gui


def get_token_expiry(token: str) -> float | None:
    """Извлекает exp из JWT. Возвращает unix timestamp или None."""
    try:
        parts = token.split(".")
        if len(parts) == 3:
            import base64, json
            payload_b64 = parts[1]
            payload_b64 += "=" * (4 - len(payload_b64) % 4)
            payload = json.loads(base64.urlsafe_b64decode(payload_b64))
            exp = payload.get("exp")
            if exp is not None:
                return float(exp)
    except Exception:
        pass
    return None


class TokenDialog:
    """Окно ручного ввода токена."""

    def __init__(self, parent, script_dir: Path, on_token_saved=None):
        self.script_dir = script_dir
        self.on_token_saved = on_token_saved

        dlg = Toplevel(parent)
        dlg.title("Настройки — Токен True API")
        dlg.geometry("520x260")
        dlg.configure(bg=COLOR_FRAME_BG)
        dlg.resizable(False, False)
        dlg.transient(parent)
        dlg.grab_set()
        self.dlg = dlg

        Label(dlg, text="Токен True API",
              font=FONT_TITLE, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(pady=(16, 4))

        Label(dlg,
              text="Войдите в ЛК markirovka.crpt.ru → F12 → Console → localStorage.getItem('token')",
              font=FONT_SMALL, bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO, wraplength=480).pack(pady=(0, 8))

        self.token_entry = Entry(dlg, font=FONT_CODE,
                                  bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                                  insertbackground=COLOR_LOG_FG,
                                  relief="flat", width=60)
        self.token_entry.pack(padx=16, pady=(0, 8), fill="x")

        current_token = os.environ.get("CHESTNYZNAK_TOKEN", "")
        if current_token:
            self.token_entry.insert(0, current_token)

        self.token_entry.focus_set()
        self.token_entry.icursor(len(current_token))

        btn_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        btn_frame.pack(pady=12)

        Button(btn_frame, text="💾 Сохранить", font=FONT_MAIN,
               bg=COLOR_BUTTON_ACCENT, fg="#1e1e2e",
               activebackground=COLOR_BUTTON_ACCENT_HOVER, activeforeground="#1e1e2e",
               relief="flat", padx=20, pady=6,
               command=self._save).pack(side="left", padx=8)

        Button(btn_frame, text="🗑 Очистить", font=FONT_MAIN,
               bg="#e64553", fg="#ffffff",
               activebackground="#d63543", activeforeground="#ffffff",
               relief="flat", padx=20, pady=6,
               command=lambda: self.token_entry.delete(0, END)).pack(side="left", padx=8)

        Button(btn_frame, text="Отмена", font=FONT_MAIN,
               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=20, pady=6,
               command=dlg.destroy).pack(side="left", padx=8)

    def _save(self):
        new_token = self.token_entry.get().strip()
        if not new_token:
            messagebox.showwarning("Пустой токен", "Введите токен или нажмите Отмена.", parent=self.dlg)
            return

        save_token_to_env(self.script_dir, new_token)

        # Сохраняем срок действия
        expires_at = get_token_expiry(new_token)
        exp_path = self.script_dir / ".token_expires"
        if expires_at is not None:
            exp_path.write_text(str(expires_at), "utf-8")
        else:
            exp_path.write_text(str(time.time() + 36000), "utf-8")

        log_to_gui("🔑 Токен сохранён в .env", "success")
        if self.on_token_saved:
            self.on_token_saved(new_token)
        self.dlg.destroy()


class UkepDialog:
    """Окно авторизации через УКЭП.

    Отображает все действующие сертификаты в таблице (№, ИНН, Subject)
    с динамической фильтрацией по ИНН.
    """

    def __init__(self, parent, script_dir: Path, on_token_saved=None):
        self.script_dir = script_dir
        self.on_token_saved = on_token_saved
        self.certs_data: list[dict] = []

        dlg = Toplevel(parent)
        dlg.title("Получить токен через УКЭП")
        dlg.geometry("720x580")
        dlg.configure(bg=COLOR_FRAME_BG)
        dlg.resizable(True, True)
        dlg.transient(parent)
        dlg.grab_set()
        self.dlg = dlg
        self.parent = parent

        # Подключаем логирование
        set_cert_log_fn(log_to_gui)
        set_jwt_log_fn(log_to_gui)

        # Заголовок
        Label(dlg, text="🔐 Авторизация через УКЭП",
              font=("Segoe UI", 13, "bold"), bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(pady=(16, 4))
        Label(dlg,
              text="Выберите сертификат из списка.\n"
                   "Токен будет получен автоматически и сохранён в настройки.",
              font=FONT_SMALL, bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO,
              wraplength=680, justify="center").pack(pady=(0, 8))

        # ── Поле фильтрации по ИНН ────────────────────────────────────
        filter_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        filter_frame.pack(fill="x", padx=16, pady=(0, 4))

        Label(filter_frame, text="🔍 Фильтр по ИНН:",
              font=FONT_MAIN, bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG
              ).pack(side="left", padx=(0, 8))

        self.inn_filter_var = StringVar()
        self.inn_filter_var.trace_add("write", self._on_inn_filter_changed)

        self.inn_filter_entry = Entry(filter_frame, font=FONT_CODE,
                                       textvariable=self.inn_filter_var,
                                       bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                                       insertbackground=COLOR_LOG_FG,
                                       relief="flat", width=20)
        self.inn_filter_entry.pack(side="left", fill="x", expand=True)

        # ── Таблица сертификатов (Treeview) ────────────────────────────
        certs_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        certs_frame.pack(fill="both", expand=True, padx=16, pady=(4, 8))

        Label(certs_frame, text="Сертификаты УКЭП:", font=FONT_MAIN,
              bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(anchor="w", pady=(0, 4))

        tree_frame = Frame(certs_frame, bg=COLOR_LOG_BG)
        tree_frame.pack(fill="both", expand=True)

        # Стиль для Treeview (Catppuccin Mocha)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Certs.Treeview",
                         background=COLOR_LOG_BG,
                         foreground=COLOR_LOG_FG,
                         fieldbackground=COLOR_LOG_BG,
                         font=FONT_CODE,
                         rowheight=24)
        style.configure("Certs.Treeview.Heading",
                         background=COLOR_HEADER_BG,
                         foreground=COLOR_BUTTON_FG,
                         font=FONT_MAIN,
                         relief="flat")
        style.map("Certs.Treeview",
                   background=[("selected", COLOR_BUTTON_ACCENT)],
                   foreground=[("selected", "#1e1e2e")])
        style.map("Certs.Treeview.Heading",
                   background=[("active", COLOR_BUTTON_ACTIVE_BG)])

        columns = ("num", "inn", "subject")
        self.certs_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            style="Certs.Treeview", selectmode="browse",
        )
        self.certs_tree.heading("num", text="№")
        self.certs_tree.heading("inn", text="ИНН")
        self.certs_tree.heading("subject", text="Сертификат")

        self.certs_tree.column("num", width=40, minwidth=30, stretch=False, anchor="center")
        self.certs_tree.column("inn", width=120, minwidth=80, stretch=False)
        self.certs_tree.column("subject", width=500, minwidth=200, stretch=True)

        vsb = Scrollbar(tree_frame, orient=VERTICAL, command=self.certs_tree.yview)
        hsb = Scrollbar(tree_frame, orient=HORIZONTAL, command=self.certs_tree.xview)
        self.certs_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.certs_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # ── Кнопки управления ──────────────────────────────────────────
        ctrl_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        ctrl_frame.pack(fill="x", padx=16, pady=(0, 4))

        Button(ctrl_frame, text="🔄 Обновить список", font=FONT_SMALL,
               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=12, pady=4,
               command=self._load_certs).pack(side="left")

        Button(ctrl_frame, text="🔍 Диагностика COM", font=FONT_SMALL,
               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=12, pady=4,
               command=self._run_diagnostics).pack(side="left", padx=(8, 0))

        self.status_var = StringVar(value="")
        Label(ctrl_frame, textvariable=self.status_var,
              font=FONT_SMALL, bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO).pack(side="right", padx=4)

        # Метод
        Label(dlg, text="Метод: JWT через УКЭП",
              font=FONT_MAIN, bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO).pack(padx=16, anchor="w", pady=(0, 8))

        # ── Кнопки действия ────────────────────────────────────────────
        btn_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        btn_frame.pack(pady=(0, 16))

        self.btn_auth = Button(btn_frame, text="🔐 Получить токен",
                                font=FONT_BUTTON_ACCENT,
                                bg=COLOR_BUTTON_ACCENT, fg="#1e1e2e",
                                activebackground=COLOR_BUTTON_ACCENT_HOVER, activeforeground="#1e1e2e",
                                relief="flat", padx=20, pady=8,
                                command=self._do_auth)
        self.btn_auth.pack(side="left", padx=8)

        Button(btn_frame, text="Отмена", font=FONT_MAIN,
               bg=COLOR_BUTTON_BG, fg=COLOR_BUTTON_FG,
               activebackground=COLOR_BUTTON_ACTIVE_BG, activeforeground=COLOR_BUTTON_FG,
               relief="flat", padx=20, pady=8,
               command=dlg.destroy).pack(side="left", padx=8)

        # Загружаем сертификаты
        self._load_certs()

    def _populate_tree(self, certs: list[dict] | None = None) -> None:
        """Заполняет Treeview сертификатами с учётом фильтра по ИНН."""
        # Очищаем
        for item in self.certs_tree.get_children():
            self.certs_tree.delete(item)

        if certs is None:
            certs = self.certs_data

        inn_filter = self.inn_filter_var.get().strip()

        num = 0
        for cert in certs:
            inn = cert.get("inn", "")
            # Динамическая фильтрация: показываем только серты,
            # ИНН которых начинается с введённого текста
            if inn_filter and not inn.startswith(inn_filter):
                continue

            num += 1
            subject = cert.get("subject", "Неизвестный")
            # Укорачиваем subject для читаемости
            display_subject = subject[:120] + "..." if len(subject) > 120 else subject

            self.certs_tree.insert("", END, iid=str(num - 1),
                                    values=(num, inn or "—", display_subject))

        if num == 0 and self.certs_data:
            self.certs_tree.insert("", END, values=("—", "—", "Нет сертификатов с таким ИНН"))

    def _on_inn_filter_changed(self, *_args) -> None:
        """Обработчик изменения поля фильтра — обновляет таблицу."""
        self._populate_tree()

    def _load_certs(self):
        """Загружает сертификаты в фоновом потоке."""
        self.certs_data.clear()
        self._populate_tree()
        self.status_var.set("⏳ Загрузка...")

        def worker():
            certs = list_all_valid_certificates()

            def on_done():
                self.status_var.set("")
                self.certs_data.clear()
                if not certs:
                    self.certs_tree.insert("", END,
                                            values=("—", "—", "Сертификаты не найдены. Нажмите «Диагностика COM»."))
                else:
                    self.certs_data.extend(certs)
                    self._populate_tree()

            self.parent.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()

    def _run_diagnostics(self):
        """Запускает диагностику COM-объектов."""
        self.status_var.set("⏳ Диагностика...")

        def worker():
            result = diagnose_com()
            log_to_gui(result, "info")

            def on_done():
                self.status_var.set("✅ См. лог")
                messagebox.showinfo("Диагностика COM", result, parent=self.dlg)

            self.parent.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()

    def _get_selected_cert_index(self) -> int | None:
        """Возвращает индекс выбранного сертификата в self.certs_data или None."""
        selection = self.certs_tree.selection()
        if not selection:
            return None

        # iid в Treeview — это строковый индекс, но при фильтрации
        # нам нужно найти реальный индекс в certs_data
        sel_values = self.certs_tree.item(selection[0], "values")
        if not sel_values:
            return None

        # Ищем по ИНН + subject для точного соответствия
        sel_inn = str(sel_values[1]) if sel_values[1] != "—" else ""
        sel_subject_start = str(sel_values[2])[:60]

        for i, cert in enumerate(self.certs_data):
            cert_inn = cert.get("inn", "")
            cert_subject_start = cert.get("subject", "")[:60]
            if cert_inn == sel_inn and cert_subject_start == sel_subject_start:
                return i

        return None

    def _do_auth(self):
        idx = self._get_selected_cert_index()

        thumbprint = ""
        inn_from_cert = ""
        cert_subject = ""

        if idx is not None:
            cert_data = self.certs_data[idx]
            thumbprint = cert_data.get("thumbprint", "")
            inn_from_cert = cert_data.get("inn", "")
            cert_subject = cert_data.get("subject", "")
        else:
            if self.certs_data:
                messagebox.showwarning(
                    "Выберите сертификат",
                    "Выберите сертификат из таблицы.",
                    parent=self.dlg,
                )
                return

        self.status_var.set("⏳ Авторизация...")
        self.btn_auth.config(state=DISABLED)

        # Логируем выбранный сертификат
        if cert_subject:
            subject_short = cert_subject[:80] + ("..." if len(cert_subject) > 80 else "")
            log_to_gui(f"📝 Выбран сертификат: {subject_short}", "info")
            if inn_from_cert:
                log_to_gui(f"   ИНН: {inn_from_cert}", "info")
            log_to_gui(f"   Thumbprint: {thumbprint}", "info")

        def worker():
            success, result = auth_jwt(thumbprint)

            def on_done():
                self.btn_auth.config(state=NORMAL)
                if success:
                    save_token_to_env(self.script_dir, result, inn_from_cert)

                    # Сохраняем срок действия
                    expires_at = get_token_expiry(result)
                    exp_path = self.script_dir / ".token_expires"
                    if expires_at is not None:
                        exp_path.write_text(str(expires_at), "utf-8")
                    else:
                        exp_path.write_text(str(time.time() + 36000), "utf-8")

                    self.status_var.set("✅ Токен получен!")

                    # Логируем каким сертификатом подписали
                    log_to_gui("═" * 50, "bold")
                    log_to_gui("🔐 JWT-токен получен через УКЭП", "success")
                    if cert_subject:
                        subject_short = cert_subject[:80] + ("..." if len(cert_subject) > 80 else "")
                        log_to_gui(f"   Подпись сертификатом: {subject_short}", "success")
                    if inn_from_cert:
                        log_to_gui(f"   ИНН: {inn_from_cert}", "success")
                    log_to_gui("═" * 50, "bold")

                    if self.on_token_saved:
                        self.on_token_saved(result)
                    messagebox.showinfo("Токен получен",
                                        "JWT-токен успешно получен и сохранён!",
                                        parent=self.dlg)
                    self.dlg.destroy()
                else:
                    self.status_var.set("❌ Ошибка")
                    messagebox.showerror("Ошибка авторизации", result, parent=self.dlg)

            self.parent.after(0, on_done)

        threading.Thread(target=worker, daemon=True).start()


class HelpDialog:
    """Окно справки."""

    def __init__(self, parent):
        from tkinter import Text as TkText

        dlg = Toplevel(parent)
        dlg.title("Инструкция")
        dlg.geometry("620x520")
        dlg.configure(bg=COLOR_FRAME_BG)
        dlg.resizable(True, True)
        dlg.transient(parent)

        text = TkText(dlg, font=FONT_MAIN, bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                      relief="flat", padx=12, pady=12, wrap="word")
        text.pack(fill="both", expand=True)

        instructions = """CISCHECKER — ПРОВЕРКА КОДОВ МАРКИРОВКИ

Как пользоваться:

1. Настройте токен True API
   Способ А: Настройки → Токен
   - Войдите в ЛК: https://markirovka.crpt.ru
   - F12 → Console → localStorage.getItem('token')
   - Скопируйте токен и вставьте в окно настроек

   Способ Б: Настройки → Получить токен через УКЭП
   - Нужен установленный КриптоПро CSP 5.x + УКЭП (RuToken)
   - Нажмите «🔄 Обновить список» → выберите сертификат
   - Используйте поле «Фильтр по ИНН» для быстрого поиска
   - Нажмите «🔐 Получить токен» — токен сохранится автоматически
   - Если сертификаты не появляются → «🔍 Диагностика COM»

2. Выберите товарную группу (выпадающий список сверху)

3. Загрузите коды маркировки (два способа):

   Способ А: Ввод вручную
   - Введите или вставьте коды в текстовое поле «Коды маркировки»
   - Каждый код на отдельной строке
   - Пустые строки будут пропущены

   Способ Б: Загрузка из файла
   - Файл → Загрузить коды из файла
   - TXT-файл, по одному коду на строку
   - Коды из файла появятся в текстовом поле

4. Нажмите «📂 Отправить список кодов»
   - Коды отправляются пачками по 100
   - Можно нажать «⏹ Стоп» для остановки

5. Выгрузите результат: «📥 Выгрузка XLSX»

Требования:
- Российский IP (серверы ЧЗ геоблокируют зарубежные)
- Python 3.8+, openpyxl
- Для УКЭП: КриптоПро CSP 5.x + pywin32 (Windows)"""

        text.insert("1.0", instructions)
        text.config(state=DISABLED)


def show_about(parent):
    """Диалог «О программе»."""
    messagebox.showinfo(
        "О программе",
        f"{APP_TITLE}\n\n"
        f"Версия: {APP_VERSION}\n"
        f"Python + Tkinter\n\n"
        f"Использует публичный и True API\n"
        f"системы маркировки «Честный Знак» (CISChecker).\n\n"
        f"Репозиторий: github.com/{GITHUB_REPO}\n"
        f"© 2026",
    )
