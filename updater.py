#!/usr/bin/env python3
"""
Обратная совместимость: реэкспорт updater-модуля.
"""
from cischecker.updater.github import (
    check_for_update,
    perform_update,
    is_frozen,
    get_current_version,
    compare_versions,
    fetch_latest_release,
    download_exe,
    create_updater_bat,
)
from cischecker import GITHUB_REPO

__all__ = [
    "check_for_update", "perform_update", "is_frozen",
    "get_current_version", "compare_versions",
    "fetch_latest_release", "download_exe", "create_updater_bat",
    "GITHUB_REPO",
]