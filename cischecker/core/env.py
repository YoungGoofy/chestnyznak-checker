"""
Управление .env файлом: загрузка и сохранение переменных окружения.
"""
from __future__ import annotations

import os
from pathlib import Path


def load_env(script_dir: Path) -> None:
    """Загружает переменные из .env, если он есть.

    Примечание: НЕ перезаписывает уже установленные переменные,
    кроме CHESTNYZNAK_TOKEN (всегда обновляется из файла).
    """
    env_path = script_dir / ".env"
    if not env_path.exists():
        return
    try:
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, value = line.partition("=")
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if not key:
                    continue
                # Токен всегда обновляем из файла (мог измениться через GUI)
                if key == "CHESTNYZNAK_TOKEN" or key not in os.environ:
                    os.environ[key] = value
    except Exception:
        pass


def save_token_to_env(script_dir: Path, token: str, inn: str = "") -> None:
    """Сохраняет токен (и опционально ИНН) в .env файл.

    Устраняет дублирование кода сохранения токена в GUI.
    """
    env_path = script_dir / ".env"
    lines: list[str] = []
    found_token = False
    found_inn = False

    if env_path.exists():
        for line in env_path.read_text("utf-8").splitlines():
            if line.startswith("CHESTNYZNAK_TOKEN="):
                lines.append(f"CHESTNYZNAK_TOKEN={token}")
                found_token = True
            elif line.startswith("CHESTNYZNAK_INN="):
                if inn:
                    lines.append(f"CHESTNYZNAK_INN={inn}")
                    found_inn = True
                # Если inn пустой — пропускаем строку (удаляем)
            else:
                lines.append(line)

    if not found_token:
        lines.append(f"CHESTNYZNAK_TOKEN={token}")
    if inn and not found_inn:
        lines.append(f"CHESTNYZNAK_INN={inn}")

    env_path.write_text("\n".join(lines) + "\n", "utf-8")

    # Обновляем os.environ
    os.environ["CHESTNYZNAK_TOKEN"] = token
    if inn:
        os.environ["CHESTNYZNAK_INN"] = inn
