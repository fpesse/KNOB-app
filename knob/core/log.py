# -*- coding: utf-8 -*-
"""Log de diagnóstico en knob_debug.log, al lado del exe (igual que v3.3)."""
import os
import time

from knob.core.paths import app_dir

LOG_PATH = os.path.join(app_dir(), "knob_debug.log")
_MAX_BYTES = 1_000_000


def _rotate():
    try:
        if os.path.getsize(LOG_PATH) > _MAX_BYTES:
            os.replace(LOG_PATH, LOG_PATH + ".old")
    except Exception:
        pass


_rotate()


def log(msg: str) -> None:
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%H:%M:%S')}] {msg}\n")
    except Exception:
        pass
