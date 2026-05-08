"""
Модуль автоматического обновления через GitHub Releases.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

from .. import APP_VERSION, GITHUB_REPO

RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_NAME = "CISChecker.exe"
CONNECT_TIMEOUT = 10
READ_TIMEOUT = 60


def get_current_version() -> str:
    return APP_VERSION


def is_frozen() -> bool:
    return getattr(sys, 'frozen', False)


def get_exe_path() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve()
    return Path(__file__).resolve().parent.parent.parent / "gui_app.py"


def fetch_latest_release() -> dict | None:
    try:
        req = Request(RELEASES_API, headers={"User-Agent": "CISChecker"})
        with urlopen(req, timeout=CONNECT_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (URLError, HTTPError, json.JSONDecodeError, OSError):
        return None

    tag = data.get("tag_name", "")
    version = tag.lstrip("v")

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
    def parse(v: str) -> tuple[int, ...]:
        return tuple(int(x) for x in v.split(".") if x.isdigit())
    cur, lat = parse(current), parse(latest)
    if lat > cur:
        return 1
    elif lat < cur:
        return -1
    return 0


def download_exe(url: str, dest: Path, progress_fn=None) -> bool:
    try:
        req = Request(url, headers={"User-Agent": "CISChecker"})
        with urlopen(req, timeout=READ_TIMEOUT) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            with open(dest, "wb") as f:
                while True:
                    chunk = resp.read(65536)
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
    bat_path = old_exe.parent / "_update.bat"
    bat_content = f"""@echo off
echo Updating CISChecker...
:wait_loop
tasklist /fi "pid eq {os.getpid()}" 2>nul | find "{os.getpid()}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait_loop
)
copy /y "{new_exe}" "{old_exe}"
if errorlevel 1 (
    echo ERROR: Failed to replace executable.
    pause
    del "%~f0"
    exit /b 1
)
del /f "{new_exe}"
start "" "{old_exe}"
del "%~f0"
exit
"""
    bat_path.write_text(bat_content, encoding="cp866")
    return bat_path


def perform_update(exe_url: str, progress_fn=None) -> tuple[bool, str]:
    if not is_frozen():
        return False, "Автообновление доступно только в .exe-режиме."

    current_exe = get_exe_path()
    if not exe_url:
        return False, "Не найден .exe в релизе на GitHub."

    temp_dir = current_exe.parent / "_update_tmp"
    temp_dir.mkdir(exist_ok=True)
    new_exe = temp_dir / EXE_NAME

    success = download_exe(exe_url, new_exe, progress_fn=lambda d, t: None)
    if not success:
        if new_exe.exists():
            new_exe.unlink()
        if temp_dir.exists():
            temp_dir.rmdir()
        return False, "Ошибка скачивания обновления."

    file_size = new_exe.stat().st_size
    if file_size < 1_000_000:
        new_exe.unlink()
        if temp_dir.exists():
            temp_dir.rmdir()
        return False, f"Скачанный файл слишком мал ({file_size} байт)."

    bat_path = create_updater_bat(current_exe, new_exe)

    try:
        subprocess.Popen(
            [str(bat_path)],
            cwd=str(current_exe.parent),
            creationflags=0x00000008,
            close_fds=True,
        )
    except OSError as e:
        return False, f"Ошибка запуска обновления: {e}"

    return True, "Обновление скачано. Приложение перезапустится..."


def check_for_update() -> tuple[bool, str, dict | None]:
    current = get_current_version()
    release = fetch_latest_release()

    if release is None:
        return False, "Не удалось проверить обновления.", None

    latest = release["version"]
    cmp = compare_versions(current, latest)

    if cmp <= 0:
        return False, f"У вас последняя версия (v{current}).", None

    notes = release["release_notes"][:500] if release["release_notes"] else ""
    msg = (
        f"Доступна новая версия v{latest} (у вас v{current}).\n\n"
        f"Что нового:\n{notes}" if notes else
        f"Доступна новая версия v{latest} (у вас v{current})."
    )

    if not release.get("exe_url"):
        msg += "\n\n⚠ .exe не найден в релизе."

    return True, msg, release
