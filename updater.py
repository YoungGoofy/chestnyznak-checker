#!/usr/bin/env python3
"""
Модуль автоматического обновления ChestnyZnakChecker.

Проверяет последний Release на GitHub, сравнивает версии,
скачивает новый .exe, подменяет старый через .bat-скрипт и перезапускает.

Работает только в .exe-режиме (PyInstaller). В режиме .py — только
проверка версии (без автозамены).
"""
from __future__ import annotations

import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# ══════════════════════════════════════════════════════════════════════
# Константы
# ══════════════════════════════════════════════════════════════════════

GITHUB_REPO = "YoungGoofy/chestnyznak-checker"
RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_NAME = "ChestnyZnakChecker.exe"

# Таймауты (секунды)
CONNECT_TIMEOUT = 10
READ_TIMEOUT = 60


def get_current_version() -> str:
    """Возвращает текущую версию приложения."""
    # Импортируем здесь, чтобы избежать циклического импорта
    try:
        from gui_app import APP_VERSION
        return APP_VERSION
    except ImportError:
        return "0.0"


def is_frozen() -> bool:
    """True, если запущен как PyInstaller .exe."""
    return getattr(sys, 'frozen', False)


def get_exe_path() -> Path:
    """Путь к текущему .exe (или .py скрипту)."""
    if is_frozen():
        return Path(sys.executable).resolve()
    return Path(__file__).resolve().parent / "gui_app.py"


def fetch_latest_release() -> dict | None:
    """
    Получает информацию о последнем Release с GitHub.

    Возвращает словарь:
        {
            "tag": "v1.2",
            "version": "1.2",
            "exe_url": "https://github.com/.../ChestnyZnakChecker.exe",
            "exe_size": 12345678,
            "release_notes": "Описание релиза",
            "published_at": "2026-05-08T...",
        }
    Или None, если не удалось получить.
    """
    try:
        req = Request(RELEASES_API, headers={"User-Agent": "ChestnyZnakChecker"})
        with urlopen(req, timeout=CONNECT_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (URLError, HTTPError, json.JSONDecodeError, OSError) as e:
        return None

    tag = data.get("tag_name", "")
    version = tag.lstrip("v")

    # Ищем .exe в ассетах релиза
    exe_url = None
    exe_size = 0
    for asset in data.get("assets", []):
        name = asset.get("name", "")
        if name.lower() == EXE_NAME.lower():
            exe_url = asset.get("browser_download_url")
            exe_size = asset.get("size", 0)
            break

    return {
        "tag": tag,
        "version": version,
        "exe_url": exe_url,
        "exe_size": exe_size,
        "release_notes": data.get("body", "") or "",
        "published_at": data.get("published_at", ""),
    }


def compare_versions(current: str, latest: str) -> int:
    """
    Сравнивает две версии вида '1.2.3'.

    Возвращает:
         1, если latest > current
         0, если равны
        -1, если latest < current
    """
    def parse(v: str) -> tuple[int, ...]:
        return tuple(int(x) for x in v.split(".") if x.isdigit())

    cur = parse(current)
    lat = parse(latest)
    if lat > cur:
        return 1
    elif lat < cur:
        return -1
    return 0


def download_exe(url: str, dest: Path, progress_fn=None) -> bool:
    """
    Скачивает .exe по URL с прогрессом.

    progress_fn(downloaded_bytes, total_bytes) — колбек прогресса.
    Возвращает True при успехе.
    """
    try:
        req = Request(url, headers={"User-Agent": "ChestnyZnakChecker"})
        with urlopen(req, timeout=READ_TIMEOUT) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            chunk_size = 65536
            downloaded = 0

            with open(dest, "wb") as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_fn and total > 0:
                        progress_fn(downloaded, total)
        return True
    except (URLError, HTTPError, OSError):
        return False


def create_updater_bat(old_exe: Path, new_exe: Path) -> Path:
    """
    Создаёт .bat-скрипт для подмены .exe:

    1. Ждёт завершения старого процесса
    2. Копирует новый .exe на место старого
    3. Удаляет временный .exe
    4. Запускает новый .exe
    5. Удаляет сам .bat
    """
    bat_path = old_exe.parent / "_update.bat"
    # Используем short name (8.3) для путей с пробелами
    old_str = str(old_exe)
    new_str = str(new_exe)

    bat_content = f"""@echo off
echo Updating ChestnyZnakChecker...
echo Waiting for old process to exit...
:wait_loop
tasklist /fi "pid eq {os.getpid()}" 2>nul | find "{os.getpid()}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait_loop
)
echo Replacing executable...
copy /y "{new_str}" "{old_str}"
if errorlevel 1 (
    echo ERROR: Failed to replace executable.
    pause
    del "%~f0"
    exit /b 1
)
del /f "{new_str}"
echo Starting new version...
start "" "{old_str}"
del "%~f0"
exit
"""
    bat_path.write_text(bat_content, encoding="cp866")
    return bat_path


def perform_update(exe_url: str, progress_fn=None) -> tuple[bool, str]:
    """
    Полный цикл обновления .exe:

    1. Скачивает новый .exe во временную папку
    2. Создаёт .bat-скрипт для подмены
    3. Запускает .bat и завершает текущий процесс

    Возвращает (success: bool, message: str).
    """
    if not is_frozen():
        return False, "Автообновление доступно только в .exe-режиме. Скачайте новую версию с GitHub вручную."

    current_exe = get_exe_path()
    if not exe_url:
        return False, "Не найден .exe в релизе на GitHub."

    # Скачиваем во временную папку рядом с .exe (не в temp, чтобы избежать проблем с правами)
    temp_dir = current_exe.parent / "_update_tmp"
    temp_dir.mkdir(exist_ok=True)
    new_exe = temp_dir / EXE_NAME

    log_msg = f"Скачиваю обновление..."
    if progress_fn:
        progress_fn(log_msg, "info")

    success = download_exe(exe_url, new_exe, progress_fn=lambda d, t: None)
    if not success:
        # Cleanup
        if new_exe.exists():
            new_exe.unlink()
        if temp_dir.exists():
            temp_dir.rmdir()
        return False, "Ошибка скачивания обновления. Проверьте подключение к интернету."

    # Проверяем размер скачанного файла
    file_size = new_exe.stat().st_size
    if file_size < 1_000_000:  # Меньше 1 МБ — подозрительно
        new_exe.unlink()
        if temp_dir.exists():
            temp_dir.rmdir()
        return False, f"Скачанный файл слишком мал ({file_size} байт). Возможно, повреждён."

    # Создаём .bat для подмены
    bat_path = create_updater_bat(current_exe, new_exe)

    # Запускаем .bat и завершаемся
    try:
        subprocess.Popen(
            [str(bat_path)],
            cwd=str(current_exe.parent),
            creationflags=0x00000008,  # DETACHED_PROCESS на Windows
            close_fds=True,
        )
    except OSError as e:
        return False, f"Ошибка запуска обновления: {e}"

    return True, "Обновление скачано. Приложение перезапустится..."


def check_for_update() -> tuple[bool, str, dict | None]:
    """
    Проверяет наличие обновления.

    Возвращает:
        (has_update: bool, message: str, release_info: dict | None)

    has_update=True — доступна новая версия.
    release_info — словарь с данными релиза (или None).
    """
    current = get_current_version()
    release = fetch_latest_release()

    if release is None:
        return False, "Не удалось проверить обновления. Проверьте подключение к интернету.", None

    latest = release["version"]
    cmp = compare_versions(current, latest)

    if cmp <= 0:
        # Текущая версия новее или равна — обновление не нужно
        # cmp = 0: версии равны; cmp = -1: текущая новее
        return False, f"У вас последняя версия (v{current}).", None

    # Доступна новая версия
    notes = release["release_notes"][:500] if release["release_notes"] else ""
    msg = (
        f"Доступна новая версия v{latest} (у вас v{current}).\n\n"
        f"Что нового:\n{notes}"
        if notes else
        f"Доступна новая версия v{latest} (у вас v{current})."
    )

    if not release.get("exe_url"):
        msg += "\n\n⚠ .exe не найден в релизе. Обновитесь вручную."

    return True, msg, release