# -*- coding: utf-8 -*-
"""config.json en %APPDATA%\\KNOB: perfiles, tema y preferencias. Mismo formato que v3.3."""
import json
import os
import threading

from knob.core.engine import NUM_CHANNELS, default_encoder_entry
from knob.core.log import log

VALID_THEMES = ("Nebula", "Cobalt", "Magma", "Onyx", "Toxic", "Sunset")
DEFAULT_THEME = "Nebula"
DEFAULT_PROFILE = "Default"
MAX_PROFILE_NAME = 24
_THEME_MIGRATION = {"Arctic": "Onyx", "Steel": "Onyx"}


def default_config_path() -> str:
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "KNOB")
    return os.path.join(base, "config.json")


def _normalize_encoders(raw) -> list:
    out = []
    for i in range(2):
        entry = default_encoder_entry()
        if isinstance(raw, list) and i < len(raw) and isinstance(raw[i], dict):
            entry.update({k: v for k, v in raw[i].items() if k in entry})
        out.append(entry)
    return out


class ProfileError(ValueError):
    pass


class Config:
    def __init__(self, path: str = None):
        self.path = path or default_config_path()
        self._lock = threading.RLock()
        self.data = self._load()
        self._migrate()   # solo en memoria: el archivo se escribe recién en el primer cambio

    # --- carga / guardado -------------------------------------------------
    def _load(self) -> dict:
        try:
            with open(self.path, encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except FileNotFoundError:
            return {}
        except Exception as e:
            log(f"config.json ilegible ({e}); se usan valores por defecto")
            return {}

    def _migrate(self) -> None:
        d = self.data
        profiles = d.get("profiles")
        if not isinstance(profiles, dict) or not profiles:
            d["profiles"] = {DEFAULT_PROFILE: {
                "channel_assignments": d.get("channel_assignments") if isinstance(d.get("channel_assignments"), dict) else {},
                "encoder_assignments": d.get("encoder_assignments", [{}, {}]),
            }}
        for pname, prof in list(d["profiles"].items()):
            if not isinstance(prof, dict):
                d["profiles"][pname] = prof = {}
            if not isinstance(prof.get("channel_assignments", {}), dict):
                prof["channel_assignments"] = {}
        if d.get("active_profile") not in d["profiles"]:
            d["active_profile"] = next(iter(d["profiles"]))
        theme = d.get("theme", DEFAULT_THEME)
        theme = _THEME_MIGRATION.get(theme, theme)
        d["theme"] = theme if theme in VALID_THEMES else DEFAULT_THEME

    def save(self) -> None:
        with self._lock:
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=True)
            os.replace(tmp, self.path)

    # --- perfil activo ----------------------------------------------------
    @property
    def active_profile(self) -> str:
        return self.data["active_profile"]

    def _profile(self) -> dict:
        return self.data["profiles"].setdefault(self.active_profile, {})

    def profile_names(self) -> list:
        return list(self.data["profiles"].keys())

    def channel_assignments(self) -> dict:
        raw = self._profile().get("channel_assignments") or {}
        return {i: raw.get(str(i)) for i in range(NUM_CHANNELS)}

    def encoder_assignments(self) -> list:
        return _normalize_encoders(self._profile().get("encoder_assignments"))

    def set_channel_assignment(self, ch: int, value) -> None:
        with self._lock:
            current = self.channel_assignments()
            current[ch] = value
            self._profile()["channel_assignments"] = {str(k): v for k, v in current.items()}
            self.save()

    def set_encoder_assignments(self, encoders: list) -> None:
        with self._lock:
            self._profile()["encoder_assignments"] = [dict(e) for e in encoders]
            self.save()

    # --- gestión de perfiles ----------------------------------------------
    def _validate_name(self, name: str, exclude: str = None) -> str:
        name = (name or "").strip()
        if not name:
            raise ProfileError("El nombre no puede estar vacío")
        if len(name) > MAX_PROFILE_NAME:
            raise ProfileError(f"El nombre puede tener hasta {MAX_PROFILE_NAME} caracteres")
        if name in self.data["profiles"] and name != exclude:
            raise ProfileError("Ya existe un perfil con ese nombre")
        return name

    def switch_profile(self, name: str) -> None:
        with self._lock:
            if name not in self.data["profiles"]:
                raise ProfileError(f"No existe el perfil {name}")
            self.data["active_profile"] = name
            self.save()

    def create_profile(self, name: str, channels: dict, encoders: list) -> str:
        with self._lock:
            name = self._validate_name(name)
            self.data["profiles"][name] = {
                "channel_assignments": {str(i): channels.get(i) for i in range(NUM_CHANNELS)},
                "encoder_assignments": [dict(e) for e in _normalize_encoders(encoders)],
            }
            self.data["active_profile"] = name
            self.save()
            return name

    def rename_profile(self, old: str, new: str) -> str:
        with self._lock:
            if old not in self.data["profiles"]:
                raise ProfileError(f"No existe el perfil {old}")
            new = self._validate_name(new, exclude=old)
            if new == old:
                return old
            self.data["profiles"] = {(new if k == old else k): v
                                     for k, v in self.data["profiles"].items()}
            if self.data["active_profile"] == old:
                self.data["active_profile"] = new
            self.save()
            return new

    def delete_profile(self, name: str) -> bool:
        with self._lock:
            profiles = self.data["profiles"]
            if len(profiles) <= 1 or name not in profiles:
                return False
            del profiles[name]
            if self.data["active_profile"] == name:
                self.data["active_profile"] = next(iter(profiles))
            self.save()
            return True

    # --- preferencias -----------------------------------------------------
    @property
    def theme(self) -> str:
        return self.data["theme"]

    @theme.setter
    def theme(self, name: str) -> None:
        if name not in VALID_THEMES:
            raise ValueError(f"Tema desconocido: {name}")
        with self._lock:
            self.data["theme"] = name
            self.save()

    @property
    def minimize_to_tray(self) -> bool:
        return bool(self.data.get("minimize_to_tray", False))

    @minimize_to_tray.setter
    def minimize_to_tray(self, value: bool) -> None:
        with self._lock:
            self.data["minimize_to_tray"] = bool(value)
            self.save()

    @property
    def last_port(self):
        return self.data.get("last_port")

    @last_port.setter
    def last_port(self, port) -> None:
        with self._lock:
            if self.data.get("last_port") == port:
                return
            self.data["last_port"] = port
            self.save()

    def exe_path(self, app_name: str):
        return self.data.get("exe_paths", {}).get(app_name.lower())

    def remember_exe_path(self, app_name: str, path: str) -> None:
        with self._lock:
            if self.exe_path(app_name) == path:
                return
            self.data.setdefault("exe_paths", {})[app_name.lower()] = path
            self.save()
