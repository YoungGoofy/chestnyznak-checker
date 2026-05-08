#!/usr/bin/env python3
"""
GUI для проверки кодов маркировки (CISChecker).
Обратная совместимость — запускает новую модульную структуру.
"""
# Экспорт APP_VERSION для обратной совместимости (updater.py legacy)
from cischecker import APP_VERSION  # noqa: F401


def main():
    from cischecker.gui.app import main as _main
    _main()


if __name__ == "__main__":
    main()
