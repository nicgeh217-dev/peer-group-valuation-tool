"""Central path resolution — works both as plain .py and as a frozen PyInstaller .exe.

When frozen, runtime files (output, cache, .env, transactions.csv, settings.json)
live next to the .exe, not in the temp extraction dir.
"""
from __future__ import annotations

import sys
from pathlib import Path


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        # Running as a PyInstaller .exe → use the folder containing the .exe
        return Path(sys.executable).parent
    # Running as a normal script
    return Path(__file__).parent


APP_DIR = app_dir()
OUTPUT_DIR = APP_DIR / "output"
CACHE_DIR = APP_DIR / "cache"
