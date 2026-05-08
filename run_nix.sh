#!/usr/bin/env bash
# CISChecker — запуск на NixOS
set -e

DIR="$(cd "$(dirname "$0")" && pwd)"

if [ ! -d "$DIR/.venv" ]; then
    echo "=== Создаю .venv ==="
    nix-shell -p python311 -p tk --run "
        cd \"$DIR\"
        python3 -m venv .venv
        source .venv/bin/activate
        pip install openpyxl python-dotenv pyinstaller
    "
fi

echo "=== Запуск GUI ==="
nix-shell -p python311 -p tk --run "
    cd \"$DIR\"
    source .venv/bin/activate
    python gui_app.py
"
