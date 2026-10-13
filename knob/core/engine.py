# -*- coding: utf-8 -*-
"""Motor del mixer: suavizado, deadband, mute, prioridad de foco y botones de encoders.

Puro: no importa Qt ni pycaw. Los efectos (volumen, teclas, salida de audio) se inyectan.
Lógica portada de app.py v3.3.0; las constantes son las mismas.
"""
import math
import time
from typing import Callable, Protocol

SMOOTH_ALPHA = 1          # velocidad del EMA (1 = salta directo al objetivo)
DEADBAND = 10             # tolerancia de ruido del ADC
MUTE_THRESHOLD = 15       # raw <= esto -> volumen 0
VOL_RATE_MS = 40          # API de volumen como máximo cada 40 ms (25 Hz)
FOCUS_MOVE_THRESHOLD = 15
DIRECT_RECLAIM_THRESHOLD = 40
OUTPUT_TURN_THRESHOLD = 35
OUTPUT_SWITCH_DEBOUNCE_MS = 800
MASTER_STEP = 0.05
VU_DB_FLOOR = -30.0
VU_PEAK_DECAY = 0.012

NUM_CHANNELS = 5
DIAL_CHANNELS = (3, 4)
ACTION_TYPES = ("click", "dblclick", "click_turn")


def raw_to_level(raw_val: float) -> float:
    """ADC (0-1023) -> nivel de audio con ley cuadrática (≈ percepción de volumen)."""
    if raw_val <= MUTE_THRESHOLD:
        return 0.0
    linear = max(0.0, min(1.0, raw_val / 1023.0))
    return linear * linear


def linear_to_vu(linear: float) -> float:
    """Pico lineal (0-1) -> altura del medidor (0-1) en escala dB con piso de -30 dB."""
    if linear <= 0:
        return 0.0
    db = 20.0 * math.log10(linear)
    return max(0.0, min(1.0, (db - VU_DB_FLOOR) / (-VU_DB_FLOOR)))


def next_peak(level: float, peak: float) -> float:
    return level if level > peak else max(0.0, peak - VU_PEAK_DECAY)


def default_encoder_entry() -> dict:
    return {"click": "none", "dblclick": "none", "click_turn": "none",
            "hotkey_keys_click": "", "hotkey_keys_dblclick": ""}


class Effects(Protocol):
    def set_volume(self, target: str, level: float) -> None: ...   # "master" | "mic" | nombre de app
    def media_key(self, action: str) -> None: ...                   # play_pause | next_track | prev_track
    def send_hotkey(self, combo: str) -> None: ...
    def step_master(self, delta: float) -> None: ...
    def cycle_output(self, direction: int) -> None: ...


class Engine:
    def __init__(self, effects: Effects, focus_provider: Callable[[], str],
                 clock: Callable[[], float] = time.monotonic):
        self.effects = effects
        self._focus = focus_provider
        self._clock = clock
        self.channel_assignments = {i: None for i in range(NUM_CHANNELS)}
        self.encoder_assignments = [default_encoder_entry(), default_encoder_entry()]
        self.smooth_vals = [512.0] * NUM_CHANNELS
        self.target_vals = [512.0] * NUM_CHANNELS
        self.dial_muted = [False, False]
        self.dial_saved_val = [512.0, 512.0]
        self.encoder_held = [False, False]
        self.hold_saved = [None, None]
        self.enc_prev_raw = [None, None]
        self._last_vol_ms = float("-inf")
        self._dev_switch_ms = float("-inf")
        self._volume_dirty = False
        # Prioridad "el último que se movió gana": app -> "direct" | "focus"
        self._channel_priority = {}
        self._last_raw_per_channel = [None] * NUM_CHANNELS
        self._focus_took_at = {}
        self._focus_prev_raw = None

    # --- asignaciones ---------------------------------------------------
    def set_assignments(self, channels: dict, encoders: list) -> None:
        self.channel_assignments = {i: channels.get(i) for i in range(NUM_CHANNELS)}
        merged = []
        for i in range(2):
            entry = default_encoder_entry()
            if i < len(encoders) and isinstance(encoders[i], dict):
                entry.update(encoders[i])
            merged.append(entry)
        self.encoder_assignments = merged

    def turn_key(self, enc_idx: int) -> str:
        return self.encoder_assignments[enc_idx].get("click_turn", "none")

    def is_muted(self, ch: int) -> bool:
        return bool(self.channel_assignments.get(ch)) and self.smooth_vals[ch] <= MUTE_THRESHOLD

    # --- hardware -------------------------------------------------------
    def on_frame(self, values: list) -> None:
        for i in range(NUM_CHANNELS):
            if i in DIAL_CHANNELS and self.dial_muted[i - 3]:
                self.target_vals[i] = 0.0      # dial muteado con el botón: ignora el hardware
                continue
            new = float(values[i])
            if abs(new - self.target_vals[i]) >= DEADBAND:
                self.target_vals[i] = new

    def tick(self) -> bool:
        """Un paso de suavizado (cada 16 ms). Devuelve True si algún valor cambió."""
        needs_update = False
        for i in range(NUM_CHANNELS):
            if i in DIAL_CHANNELS:
                e = i - 3
                if self.dial_muted[e]:
                    if self.smooth_vals[i] != 0.0:
                        self.smooth_vals[i] = 0.0
                        needs_update = True
                    continue
                if self.encoder_held[e] and self.turn_key(e) != "none":
                    if self.hold_saved[e] is not None:
                        self.smooth_vals[i] = self.hold_saved[e]   # congela UI/audio, no el target
                    needs_update = True
                    continue
            diff = self.target_vals[i] - self.smooth_vals[i]
            if abs(diff) < 0.5:
                if self.smooth_vals[i] != self.target_vals[i]:
                    self.smooth_vals[i] = self.target_vals[i]
                    needs_update = True
            else:
                self.smooth_vals[i] += diff * SMOOTH_ALPHA
                needs_update = True
        if needs_update:
            self._volume_dirty = True
        if self._volume_dirty:
            self._apply_volumes()
        return needs_update

    def _now_ms(self) -> float:
        return self._clock() * 1000.0

    def _apply_volumes(self) -> None:
        now = self._now_ms()
        if now - self._last_vol_ms < VOL_RATE_MS:
            return                                  # queda pendiente para el próximo tick
        self._last_vol_ms = now
        self._volume_dirty = False
        for i in range(NUM_CHANNELS):
            if i in DIAL_CHANNELS:
                e = i - 3
                if self.encoder_held[e]:
                    turn = self.turn_key(e)
                    if turn == "output_device":
                        curr = self.target_vals[i]  # valor real del hardware, no el congelado
                        prev = self.enc_prev_raw[e]
                        if prev is not None and abs(curr - prev) > OUTPUT_TURN_THRESHOLD:
                            self._cycle_output(1 if curr > prev else -1)
                        self.enc_prev_raw[e] = curr
                        continue
                    self.enc_prev_raw[e] = None
                    if turn != "none":
                        self.apply_key(turn, self.smooth_vals[i])
                        continue
            self.apply_channel(i, self.smooth_vals[i])

    # --- volumen --------------------------------------------------------
    def apply_key(self, key: str, raw_val: float) -> None:
        """Aplica un valor a un destino directo (master/mic/focus/app), sin canal."""
        level = raw_to_level(raw_val)
        if key == "focus":
            self._set_focus_volume(level)
        else:
            self.effects.set_volume(key, level)

    def apply_channel(self, ch: int, raw_val: float) -> None:
        assignment = self.channel_assignments.get(ch)
        if not assignment:
            return
        level = raw_to_level(raw_val)
        if assignment in ("master", "mic"):
            self.effects.set_volume(assignment, level)
            return
        if assignment == "focus":
            self._set_focus_volume(level, raw_val)
            return
        self._last_raw_per_channel[ch] = raw_val
        if self._channel_priority.get(assignment, "direct") == "focus":
            ref = self._focus_took_at.get(assignment)
            if ref is not None and abs(raw_val - ref) > DIRECT_RECLAIM_THRESHOLD:
                self._channel_priority[assignment] = "direct"
                self._focus_took_at.pop(assignment, None)
        if self._channel_priority.get(assignment, "direct") == "direct":
            self.effects.set_volume(assignment, level)

    def _set_focus_volume(self, level: float, raw_val=None) -> None:
        exe = self._focus()
        if not exe:
            self._focus_prev_raw = raw_val
            return
        moving = raw_val is not None and (
            self._focus_prev_raw is None or abs(raw_val - self._focus_prev_raw) > FOCUS_MOVE_THRESHOLD)
        self._focus_prev_raw = raw_val
        if moving and self._channel_priority.get(exe) != "focus":
            for ch, app in self.channel_assignments.items():
                if app == exe and self._last_raw_per_channel[ch] is not None:
                    self._focus_took_at[exe] = self._last_raw_per_channel[ch]
                    break
            self._channel_priority[exe] = "focus"
        if self._channel_priority.get(exe, "direct") == "focus":
            self.effects.set_volume(exe, level)

    def _cycle_output(self, direction: int) -> None:
        now = self._now_ms()
        if now - self._dev_switch_ms < OUTPUT_SWITCH_DEBOUNCE_MS:
            return
        self._dev_switch_ms = now
        self.effects.cycle_output(direction)
