# -*- coding: utf-8 -*-
"""Íconos reales de las apps desde su .exe (QFileIconProvider da alta resolución y alfa correcto)."""
import os
import time

from PySide6.QtCore import QFileInfo, QSize
from PySide6.QtWidgets import QFileIconProvider

RETRY_S = 10.0     # si no se encontró el exe, reintentar como mucho cada 10 s (igual que v3.3)


class AppIcons:
    def __init__(self, exe_lookup):
        self._lookup = exe_lookup        # app_name -> ruta del exe o None
        self._provider = QFileIconProvider()
        self._cache = {}
        self._misses = {}

    def icon_for(self, app_name: str):
        key = app_name.lower()
        if key in self._cache:
            return self._cache[key]
        last_miss = self._misses.get(key)
        if last_miss is not None and time.monotonic() - last_miss < RETRY_S:
            return None
        path = self._lookup(app_name)
        if path and os.path.exists(path):
            icon = self._provider.icon(QFileInfo(path))
            if not icon.isNull():
                pixmap = icon.pixmap(QSize(64, 64))
                if not pixmap.isNull():
                    self._cache[key] = pixmap
                    self._misses.pop(key, None)
                    return pixmap
        self._misses[key] = time.monotonic()
        return None
