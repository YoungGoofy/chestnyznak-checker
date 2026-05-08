"""
Потокобезопасный виджет логов с цветными тегами.
"""
from __future__ import annotations

import queue
from tkinter import Text, Scrollbar, Frame, VERTICAL, END, DISABLED, NORMAL

from .theme import (
    COLOR_LOG_BG, COLOR_LOG_FG, COLOR_LOG_ERROR,
    COLOR_LOG_SUCCESS, COLOR_LOG_INFO, COLOR_LOG_WARN,
    FONT_CODE, FONT_CODE_BOLD, LOG_MAX_LINES,
)

# ── Глобальная очередь логов ───────────────────────────────────────────
log_queue: queue.Queue = queue.Queue()


def log_to_gui(message: str, tag: str = "info") -> None:
    """Кладёт сообщение в очередь для отображения в GUI."""
    log_queue.put((message, tag))


class LogWidget:
    """Цветное окно логов с автопрокруткой."""

    def __init__(self, parent: Frame):
        self.frame = Frame(parent, bg=COLOR_LOG_BG)

        self.text = Text(
            self.frame,
            font=FONT_CODE,
            bg=COLOR_LOG_BG, fg=COLOR_LOG_FG,
            insertbackground=COLOR_LOG_FG,
            relief="flat",
            padx=8, pady=8,
            wrap="word",
            state=DISABLED,
        )
        self.text.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(self.frame, orient=VERTICAL, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.config(yscrollcommand=scrollbar.set)

        # Теги
        self.text.tag_config("info", foreground=COLOR_LOG_INFO)
        self.text.tag_config("success", foreground=COLOR_LOG_SUCCESS)
        self.text.tag_config("error", foreground=COLOR_LOG_ERROR)
        self.text.tag_config("warn", foreground=COLOR_LOG_WARN)
        self.text.tag_config("bold", font=FONT_CODE_BOLD)

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)

    def append(self, message: str, tag: str = "info") -> None:
        """Добавляет строку в лог."""
        self.text.config(state=NORMAL)
        if self.text.index("end-1c") != "1.0":
            self.text.insert(END, "\n")
        self.text.insert(END, message, tag)
        self.text.see(END)

        line_count = int(self.text.index("end-1c").split(".")[0])
        if line_count > LOG_MAX_LINES:
            self.text.delete("1.0", f"{line_count - LOG_MAX_LINES}.0")
        self.text.config(state=DISABLED)

    def poll_queue(self, root) -> None:
        """Периодически забирает сообщения из очереди."""
        try:
            while True:
                msg, tag = log_queue.get_nowait()
                self.append(msg, tag)
        except queue.Empty:
            pass
        root.after(100, self.poll_queue, root)
