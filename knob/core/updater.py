# -*- coding: utf-8 -*-
"""Versión de la app y actualización desde GitHub Releases."""
import json
import os
import tempfile
import urllib.request

from knob.core.log import log
from knob.core.paths import app_dir, resource_path

RELEASES_URL = "https://api.github.com/repos/fpesse/KNOB-app/releases/latest"


def parse_version(v: str) -> tuple:
    try:
        return tuple(int(x) for x in v.strip().lstrip("v").split("."))
    except Exception:
        return (0,)


def read_version() -> str:
    # Primero junto al exe (donde lo copia el instalador), después el empaquetado.
    for path in (os.path.join(app_dir(), "version.txt"), resource_path("version.txt")):
        try:
            with open(path, encoding="utf-8") as f:
                v = f.read().strip()
            if v:
                return v
        except Exception:
            continue
    return "v?.?.?"


def check_for_update(current: str):
    """Devuelve (tag, url_del_instalador) si hay una versión más nueva; si no, None."""
    try:
        req = urllib.request.Request(RELEASES_URL, headers={
            "User-Agent": "KNOB-App", "Accept": "application/vnd.github+json"})
        with urllib.request.urlopen(req, timeout=8) as r:
            data = json.loads(r.read())
        tag = data.get("tag_name", "")
        url = next((a["browser_download_url"] for a in data.get("assets", [])
                    if a.get("name", "").endswith(".exe")), None)
        if url and parse_version(tag) > parse_version(current):
            return tag, url
    except Exception as e:
        log(f"check_for_update: {e}")
    return None


def download(url: str, on_progress) -> str:
    """Descarga el instalador a un temporal y devuelve la ruta. Lanza excepción si falla."""
    fd, path = tempfile.mkstemp(suffix=".exe", prefix="KNOB_update_")
    os.close(fd)
    with urllib.request.urlopen(url, timeout=60) as r:
        total = int(r.headers.get("Content-Length", 0))
        done = 0
        with open(path, "wb") as f:
            while True:
                buf = r.read(65536)
                if not buf:
                    break
                f.write(buf)
                done += len(buf)
                if total:
                    on_progress(int(done * 100 / total))
    return path


def launch_installer(path: str) -> None:
    os.startfile(path)
