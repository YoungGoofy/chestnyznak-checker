"""
Диалоговые окна: Токен, УКЭП, Справка, О программе.
"""
from __future__ import annotations

import os
import threading
import time
from pathlib import Path
from tkinter import (
    Toplevel, Frame, Label, Button, Entry, Listbox, Scrollbar,
    StringVar, END, VERTICAL, DISABLED, NORMAL, messagebox,
)

from .. import APP_VERSION, APP_TITLE, GITHUB_REPO
from ..core.env import save_token_to_env
from ..auth.certificates import list_certificates, diagnose_com, set_log_fn as set_cert_log_fn
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
    """Окно авторизации через УКЭП."""

    def __init__(self, parent, script_dir: Path, on_token_saved=None):
        self.script_dir = script_dir
        self.on_token_saved = on_token_saved
        self.certs_data: list[dict] = []

        dlg = Toplevel(parent)
        dlg.title("Получить токен через УКЭП")
        dlg.geometry("600x550")
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
              text="Подключите USB-токен (RuToken) и выберите сертификат.\n"
                   "Токен будет получен автоматически и сохранён в настройки.",
              font=FONT_SMALL, bg=COLOR_FRAME_BG, fg=COLOR_LOG_INFO,
              wraplength=560, justify="center").pack(pady=(0, 12))

        # Список сертификатов
        certs_frame = Frame(dlg, bg=COLOR_FRAME_BG)
        certs_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        Label(certs_frame, text="Сертификаты УКЭП:", font=FONT_MAIN,
              bg=COLOR_FRAME_BG, fg=COLOR_BUTTON_FG).pack(anchor="w", pady=(0, 4))

        lb_frame = Frame(certs_frame, bg=COLOR_LOG_BG)
        lb_frame.pack(fill="both", expand=True)

        self.certs_listbox = Listbox(lb_frame, font=FONT_CODE,
                                      bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
                                      selectbackground=COLOR_BUTTON_ACCENT,
                                      selectforeground="#1e1e2e",
                                      relief="flat", height=6)
        sb = Scrollbar(lb_frame, orient=VERTICAL, command=self.certs_listbox.yview)
        self.certs_listbox.config(yscrollcommand=sb.set)
        self.certs_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Кнопки управления
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

        # Кнопки действия
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

    def _load_certs(self):
        self.certs_listbox.delete(0, END)
        self.certs_data.clear()
        self.status_var.set("⏳ Загрузка...")

        def worker():
            certs = list_certificates()

            def on_done():
                self.status_var.set("")
                if not certs:
                    self.certs_listbox.insert(END, "  Сертификаты не найдены.")
                    self.certs_listbox.insert(END, "  Нажмите «Диагностика COM» для проверки.")
                else:
                    for cert in certs:
                        subject = cert.get("subject", "Неизвестный")
                        inn = cert.get("inn", "")
                        display = f"  {subject}"
                        if inn:
                            display += f" (ИНН: {inn})"
                        self.certs_listbox.insert(END, display)
                        self.certs_data.append(cert)

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

    def _do_auth(self):
        thumbprint = ""
        sel = self.certs_listbox.curselection()
        if sel and sel[0] < len(self.certs_data):
            thumbprint = self.certs_data[sel[0]].get("thumbprint", "")

        self.status_var.set("⏳ Авторизация...")
        self.btn_auth.config(state=DISABLED)

        inn_from_cert = ""
        if sel and sel[0] < len(self.certs_data):
            inn_from_cert = self.certs_data[sel[0]].get("inn", "")

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
                    log_to_gui("🔐 JWT-токен получен через УКЭП", "success")
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
        dlg.geometry("600x450")
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
   - Нажмите «🔐 Получить токен» — токен сохранится автоматически
   - Если сертификаты не появляются → «🔍 Диагностика COM»

2. Выберите товарную группу (выпадающий список сверху)

3. Загрузите список кодов (Файл → Загрузить коды)
   - TXT-файл, по одному коду на строку

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
