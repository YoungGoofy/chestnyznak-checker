"""
UI автообновления: кнопка-индикатор + проверка через GitHub Releases.
"""
from __future__ import annotations

import threading
from tkinter import Button, messagebox, DISABLED, NORMAL

from .. import APP_VERSION
from ..updater.github import (
    check_for_update as _check_update,
    perform_update as _perform_update,
    is_frozen as _is_frozen,
)
from .theme import *
from .log_widget import log_to_gui


class UpdaterUI:
    """Управляет кнопкой-индикатором обновлений."""

    def __init__(self, parent_frame, root):
        self.root = root
        self._pending_update: dict | None = None

        self.btn = Button(parent_frame, text="🔄 Проверка...",
                          font=FONT_SMALL,
                          bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE,
                          activebackground=COLOR_UPDATE_NONE, activeforeground=COLOR_UPDATE_FG_NONE,
                          relief="flat", padx=10, pady=2,
                          state=DISABLED,
                          command=self._on_click)
        self.btn.pack(side="right", padx=(0, 8), pady=12)

    def auto_check(self) -> None:
        """Тихая проверка при старте."""
        log_to_gui("🔄 Проверка обновлений...", "info")
        self._set_checking()

        def worker():
            has_update, message, release_info = _check_update()

            def on_result():
                self._pending_update = release_info if has_update else None
                if has_update:
                    ver = release_info.get("version", "?") if release_info else "?"
                    log_to_gui(f"🟢 Доступно обновление: v{ver} (у вас v{APP_VERSION})", "success")
                    self._set_available(ver)
                else:
                    log_to_gui("✓ Установлена последняя версия.", "info")
                    self._set_none()

            self.root.after(0, on_result)

        threading.Thread(target=worker, daemon=True).start()

    def manual_check(self) -> None:
        """Ручная проверка из меню."""
        log_to_gui("🔄 Проверяю обновления...", "info")
        self._set_checking()

        def worker():
            has_update, message, release_info = _check_update()

            def on_result():
                self._pending_update = release_info if has_update else None
                if has_update and release_info:
                    ver = release_info.get("version", "?")
                    log_to_gui(f"🟢 Доступно обновление: v{ver} (у вас v{APP_VERSION})", "success")
                    self._set_available(ver)
                    result = messagebox.askyesno(
                        "Доступно обновление",
                        f"Доступна новая версия v{ver} (у вас v{APP_VERSION}).\n\n"
                        f"Скачать и установить обновление?\n\n"
                        f"{'⚠ Приложение перезапустится.' if _is_frozen() else 'ℹ Автообновление доступно только в .exe-режиме.'}",
                    )
                    if result:
                        self._apply_update(release_info)
                else:
                    log_to_gui("✓ Установлена последняя версия.", "info")
                    self._set_none()
                    messagebox.showinfo("Обновление", message)

            self.root.after(0, on_result)

        threading.Thread(target=worker, daemon=True).start()

    def _on_click(self) -> None:
        if self._pending_update:
            ver = self._pending_update.get("version", "?")
            result = messagebox.askyesno(
                "Обновление",
                f"Скачать и установить v{ver}?\n\n"
                f"{'⚠ Приложение перезапустится.' if _is_frozen() else 'ℹ Автообновление доступно только в .exe-режиме.'}",
            )
            if result:
                self._apply_update(self._pending_update)

    def _apply_update(self, release_info: dict) -> None:
        from .. import GITHUB_REPO
        exe_url = release_info.get("exe_url")
        tag = release_info.get("tag", "")
        url = f"https://github.com/{GITHUB_REPO}/releases/tag/{tag}"

        if not exe_url or not _is_frozen():
            import webbrowser
            reason = ".exe не найден в релизе." if not exe_url else "Автообновление работает только в .exe."
            messagebox.showinfo("Обновление", f"{reason}\n\nСкачайте с GitHub:\n{url}")
            webbrowser.open(url)
            return

        log_to_gui("📥 Скачиваю обновление...", "info")

        def worker():
            success, msg = _perform_update(exe_url, progress_fn=lambda d, t: None)

            def on_result():
                if success:
                    log_to_gui("✅ Обновление скачано. Перезапускаю...", "bold")
                    self.root.after(1000, self.root.destroy)
                else:
                    log_to_gui(f"❌ {msg}", "error")
                    messagebox.showerror("Ошибка обновления", msg)

            self.root.after(0, on_result)

        threading.Thread(target=worker, daemon=True).start()

    def _set_checking(self):
        self.btn.config(text="🔄 Проверка...", state=DISABLED,
                        bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE)

    def _set_none(self):
        self.btn.config(text="✓ Обновлений нет", state=DISABLED,
                        bg=COLOR_UPDATE_NONE, fg=COLOR_UPDATE_FG_NONE,
                        activebackground=COLOR_UPDATE_NONE, activeforeground=COLOR_UPDATE_FG_NONE)

    def _set_available(self, ver: str):
        self.btn.config(text=f"⬆ Обновление v{ver}", state=NORMAL,
                        bg=COLOR_UPDATE_AVAIL, fg=COLOR_UPDATE_FG_AVAIL,
                        activebackground="#8be89c", activeforeground=COLOR_UPDATE_FG_AVAIL)
