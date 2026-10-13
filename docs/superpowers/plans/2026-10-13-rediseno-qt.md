# Rediseño KNOB App en PySide6: plan de implementación

> **Para agentes:** SUB-SKILL REQUERIDA: usar superpowers:subagent-driven-development (recomendado) o superpowers:executing-plans para implementar este plan tarea por tarea. Los pasos usan checkboxes (`- [ ]`).

**Objetivo:** reemplazar la interfaz customtkinter de `app.py` por una interfaz PySide6 estilo "Glow" (v4.0.0), conservando todo el comportamiento de hardware y audio de v3.3.0.

**Arquitectura:** `knob/core/` tiene la lógica sin interfaz. `protocol.py` y `engine.py` son puros y se testean con pytest. `serial_link`, `audio` y `system` envuelven el hardware y Windows. `knob/ui/` tiene widgets pintados con QPainter y superposiciones dentro de la ventana. `knob/controller.py` une todo, y `main.py` es la entrada.

**Stack:** Python 3.11 (venv en `E:\KNOB\venv`), PySide6-Essentials 6.8.1, pycaw 20251023, comtypes 1.4.16, psutil 7.2.2, pyserial 3.5, pytest 8.3.4, PyInstaller 6.19.0, Inno Setup 6.

**Spec:** `docs/superpowers/specs/2026-10-13-rediseno-qt-design.md`

## Restricciones globales

- Rama de trabajo: `rediseno-qt`. No tocar `master`.
- `app.py` no se modifica. Se borra recién en la Tarea 13, con el OK del usuario.
- Protocolo serial: 115200 baudios, `{KNOB_CMD_AWAIT_INIT}` / `{KNOB_CMD_START}\n` / `{KNOB_LOG_INITIALIZED}`, frame `v0|v1|v2|v3|v4|btnL|btnR`, botones `0` ninguno, `1` click, `2` doble click, `3` mantener, `4` fin de mantener.
- Constantes del motor (copiadas de v3.3): `SMOOTH_ALPHA=1`, `DEADBAND=10`, `MUTE_THRESHOLD=15`, `VOL_RATE_MS=40`, foco se mueve si >15, el canal directo recupera si >40, giro de salida >35, antirrebote de salida 800 ms, paso de master 0.05, VU piso -30 dB y caída de pico 0.012.
- Config: `%APPDATA%\KNOB\config.json`, mismo formato que v3.3. Temas válidos: Nebula, Cobalt, Magma, Onyx, Toxic y Sunset. Arctic y Steel migran a Onyx; cualquier otro nombre, a Nebula.
- Ventana de 520×620 fija. Textos de interfaz, comentarios y commits en español. Los archivos llevan `# -*- coding: utf-8 -*-`.
- Comandos desde `E:\KNOB` en bash: Python = `venv/Scripts/python.exe`, tests = `venv/Scripts/python.exe -m pytest`.
- Commits con la línea final `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Desvíos y correcciones respecto de v3.3 (documentados a propósito)

1. **Íconos de apps:** se obtienen con `QFileIconProvider` en `knob/ui/icons.py` (alta resolución y alfa correcto) en lugar de `core/icons.py` con GDI.
2. **Frames seriales:** el hilo guarda el último frame y el timer de 16 ms lo toma con `take_latest()`, igual que v3.3 (`latest_vals`). No se emite una señal por frame, porque el firmware manda unos 400 frames por segundo.
3. **Volumen pendiente:** en v3.3, si el último movimiento caía dentro de la ventana de 40 ms, ese valor final nunca se aplicaba. El motor ahora recuerda que quedó un cambio pendiente y lo aplica en el siguiente tick permitido.
4. **Medidor de "App en foco":** en v3.3 siempre marcaba 0, porque llamaba `GetPeakValue` sobre una lista. Se corrige.
5. **"Mantener y girar" hacia master, mic o foco:** se conserva exactamente el comportamiento de v3.3. Mientras está mantenido, el destino recibe el valor congelado de la perilla. Un test lo documenta. Corregirlo es una decisión aparte del usuario (ver la nota al final del plan).

## Mapa de archivos

```
requirements.txt, pytest.ini
main.py                          entrada (instancia única, QApplication, Controller)
knob/__init__.py
knob/controller.py               une core + ui
knob/core/__init__.py
knob/core/paths.py               app_dir, resource_path
knob/core/log.py                 log() a knob_debug.log
knob/core/protocol.py            parse_frame, ButtonEdgeDetector (puro)
knob/core/engine.py              Engine, raw_to_level, linear_to_vu, next_peak (puro)
knob/core/config.py              Config, ProfileError
knob/core/system.py              instancia única, inicio con Windows, teclas, FocusTracker, barra de título
knob/core/updater.py             versión y GitHub Releases
knob/core/audio.py               AudioBackend, OutputDevices, find_exe_path
knob/core/serial_link.py         SerialLink(QObject), handshake, candidate_ports
knob/ui/__init__.py
knob/ui/theme.py                 Theme, THEMES, fuentes y QSS
knob/ui/icons.py                 AppIcons
knob/ui/main_window.py           MainWindow
knob/ui/preview.py               vista previa con datos falsos (+ --screenshot)
knob/ui/tray.py                  Tray
knob/ui/widgets/__init__.py
knob/ui/widgets/glass.py         GlassCard, Backdrop
knob/ui/widgets/fader.py         Fader
knob/ui/widgets/vu_meter.py      VuMeter
knob/ui/widgets/knob_dial.py     KnobDial
knob/ui/widgets/brand.py         BrandLabel
knob/ui/widgets/status_pill.py   StatusPill
knob/ui/widgets/toggle.py        Toggle
knob/ui/widgets/icon_badge.py    IconBadge
knob/ui/widgets/channel_card.py  ChannelCard
knob/ui/widgets/encoder_card.py  EncoderCard
knob/ui/widgets/output_bar.py    OutputBar
knob/ui/overlays/__init__.py
knob/ui/overlays/base.py         OverlayHost, Panel, section_label, OptionRow
knob/ui/overlays/assign_channel.py
knob/ui/overlays/encoder_config.py   (+ build_combo para atajos)
knob/ui/overlays/profile_name.py
knob/ui/overlays/settings_panel.py
tests/__init__.py, tests/conftest.py, tests/fakes.py
tests/test_protocol.py, test_engine.py, test_engine_buttons.py, test_config.py,
tests/test_system.py, test_audio.py, test_serial_link.py, test_widgets.py, test_main_window.py,
tests/test_overlays.py, test_imports.py
```

---

### Tarea 1: Andamiaje + protocolo serial

**Archivos:**
- Crear: `requirements.txt`, `pytest.ini`, `knob/__init__.py`, `knob/core/__init__.py`, `knob/ui/__init__.py`, `knob/ui/widgets/__init__.py`, `knob/ui/overlays/__init__.py`, `tests/__init__.py`, `tests/conftest.py`, `knob/core/protocol.py`
- Test: `tests/test_protocol.py`

**Interfaces:**
- Produce: `protocol.BAUD_RATE`, `CMD_AWAIT_INIT: str`, `CMD_START: bytes`, `LOG_INITIALIZED: str`, `BTN_NONE..BTN_HELD_END` (0..4), `Frame(values: list[int], buttons: list[int|None])`, `parse_frame(line: str) -> Frame | None`, `is_stream_line(line: str) -> bool`, `ButtonEdgeDetector.feed(buttons) -> list[tuple[int, int]]`. Fixture `qapp` en conftest.

- [ ] **Paso 1: Dependencias y estructura**

`requirements.txt`:
```
pyserial==3.5
pycaw==20251023
comtypes==1.4.16
psutil==7.2.2
PySide6-Essentials==6.8.1
pyinstaller==6.19.0
pytest==8.3.4
```

`pytest.ini`:
```ini
[pytest]
testpaths = tests
pythonpath = .
addopts = -q
```

Los `__init__.py` de `knob/`, `knob/core/`, `knob/ui/`, `knob/ui/widgets/`, `knob/ui/overlays/` y `tests/` contienen solo:
```python
# -*- coding: utf-8 -*-
```

`tests/conftest.py`:
```python
# -*- coding: utf-8 -*-
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    yield app
```

Instalar:
```bash
venv/Scripts/python.exe -m pip install PySide6-Essentials==6.8.1 pytest==8.3.4
```
Esperado: `Successfully installed ...` (o "already satisfied").

- [ ] **Paso 2: Escribir el test que falla**

`tests/test_protocol.py`:
```python
# -*- coding: utf-8 -*-
from knob.core import protocol
from knob.core.protocol import ButtonEdgeDetector, parse_frame


def test_frame_de_siete_valores():
    f = parse_frame("10|20|30|40|50|1|0\r\n")
    assert f.values == [10, 20, 30, 40, 50]
    assert f.buttons == [1, 0]


def test_frame_legado_de_cinco_valores():
    f = parse_frame("1|2|3|4|5")
    assert f.values == [1, 2, 3, 4, 5]
    assert f.buttons == [None, None]


def test_frames_invalidos():
    assert parse_frame("") is None
    assert parse_frame("1|2|3|4") is None
    assert parse_frame("{KNOB_LOG_INITIALIZED}") is None
    assert parse_frame("a|2|3|4|5") is None


def test_boton_no_numerico_queda_none():
    assert parse_frame("1|2|3|4|5|x|2").buttons == [None, 2]


def test_is_stream_line():
    assert protocol.is_stream_line("1|2|3|4|5")
    assert not protocol.is_stream_line(protocol.CMD_AWAIT_INIT)


def test_detector_dispara_solo_en_flanco():
    d = ButtonEdgeDetector()
    assert d.feed([0, 0]) == []
    assert d.feed([1, 0]) == [(0, protocol.BTN_CLICK)]
    assert d.feed([1, 0]) == []
    assert d.feed([0, 0]) == []
    assert d.feed([3, 3]) == [(0, 3), (1, 3)]
    assert d.feed([3, 3]) == []          # el firmware repite HELD mientras se mantiene
    assert d.feed([4, 0]) == [(0, 4)]
    assert d.feed([None, None]) == []
```

- [ ] **Paso 3: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_protocol.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.core.protocol'`.

- [ ] **Paso 4: Implementar**

`knob/core/protocol.py`:
```python
# -*- coding: utf-8 -*-
"""Protocolo serial KNOB <-> app. Es el contrato con el firmware: no cambiar sin tocar ambos lados."""
from dataclasses import dataclass

BAUD_RATE = 115200
CMD_AWAIT_INIT = "{KNOB_CMD_AWAIT_INIT}"
CMD_START = b"{KNOB_CMD_START}\n"
LOG_INITIALIZED = "{KNOB_LOG_INITIALIZED}"
NUM_CHANNELS = 5

BTN_NONE = 0
BTN_CLICK = 1
BTN_DBLCLICK = 2
BTN_HELD_START = 3
BTN_HELD_END = 4


@dataclass
class Frame:
    values: list    # 5 valores analógicos 0-1023
    buttons: list   # [btnL, btnR]; None si el frame no los trae o no son numéricos


def parse_frame(line: str):
    """Parsea 'v0|v1|v2|v3|v4|btnL|btnR'. Acepta >= 5 valores (frames legados)."""
    partes = line.strip().split("|")
    if len(partes) < NUM_CHANNELS:
        return None
    try:
        values = [int(x) for x in partes[:NUM_CHANNELS]]
    except ValueError:
        return None
    buttons = []
    for col in (5, 6):
        if len(partes) > col:
            try:
                buttons.append(int(partes[col]))
            except ValueError:
                buttons.append(None)
        else:
            buttons.append(None)
    return Frame(values, buttons)


def is_stream_line(line: str) -> bool:
    return line.count("|") >= 4


class ButtonEdgeDetector:
    """El firmware repite el código mientras dura el evento; solo se dispara cuando cambia."""

    def __init__(self):
        self._prev = [BTN_NONE, BTN_NONE]

    def feed(self, buttons):
        events = []
        for idx, ev in enumerate(buttons[:2]):
            if ev is None:
                continue
            if ev != self._prev[idx]:
                self._prev[idx] = ev
                if ev > 0:
                    events.append((idx, ev))
        return events
```

- [ ] **Paso 5: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_protocol.py`
Esperado: `6 passed`.

- [ ] **Paso 6: Commit**

```bash
git add requirements.txt pytest.ini knob tests
git commit -m "feat: andamiaje del paquete knob y parser del protocolo serial" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 2: Motor: suavizado, volumen y prioridad de foco

**Archivos:**
- Crear: `knob/core/engine.py`, `tests/fakes.py`
- Test: `tests/test_engine.py`

**Interfaces:**
- Consume: `knob.core.protocol` (constantes `BTN_*`, recién en la Tarea 3).
- Produce:
  - Constantes del motor (ver Restricciones globales), `NUM_CHANNELS=5`, `DIAL_CHANNELS=(3, 4)`, `ACTION_TYPES=("click", "dblclick", "click_turn")`.
  - `raw_to_level(raw: float) -> float`, `linear_to_vu(linear: float) -> float`, `next_peak(level: float, peak: float) -> float`, `default_encoder_entry() -> dict`.
  - Protocolo `Effects` con `set_volume(target: str, level: float)`, `media_key(action: str)`, `send_hotkey(combo: str)`, `step_master(delta: float)` y `cycle_output(direction: int)`.
  - `Engine(effects, focus_provider: Callable[[], str], clock=time.monotonic)` con:
    - Atributos `channel_assignments: dict[int, str|None]`, `encoder_assignments: list[dict]`, `smooth_vals`, `target_vals`, `dial_muted`, `encoder_held` y `hold_saved`.
    - Métodos `set_assignments(channels, encoders)`, `turn_key(enc_idx) -> str`, `on_frame(values: list[int])`, `tick() -> bool`, `apply_channel(ch, raw)`, `apply_key(key, raw)`, `is_muted(ch) -> bool` y `on_button(enc_idx, code)` (este último en la Tarea 3).
  - `tests/fakes.py`: `FakeEffects` (lista `calls`) y `FakeClock` (atributo `t` en segundos).

- [ ] **Paso 1: Fakes y test que falla**

`tests/fakes.py`:
```python
# -*- coding: utf-8 -*-
"""Dobles de prueba para el motor."""


class FakeEffects:
    def __init__(self):
        self.calls = []

    def set_volume(self, target, level):
        self.calls.append(("set_volume", target, round(level, 4)))

    def media_key(self, action):
        self.calls.append(("media_key", action))

    def send_hotkey(self, combo):
        self.calls.append(("send_hotkey", combo))

    def step_master(self, delta):
        self.calls.append(("step_master", delta))

    def cycle_output(self, direction):
        self.calls.append(("cycle_output", direction))

    def volumes(self):
        return [c for c in self.calls if c[0] == "set_volume"]


class FakeClock:
    def __init__(self, t=100.0):
        self.t = t

    def __call__(self):
        return self.t
```

`tests/test_engine.py`:
```python
# -*- coding: utf-8 -*-
import pytest

from knob.core.engine import (DEADBAND, Engine, MUTE_THRESHOLD, linear_to_vu, next_peak,
                              raw_to_level)
from tests.fakes import FakeClock, FakeEffects


def lv(raw):
    return round(raw_to_level(raw), 4)


def make(focus=""):
    fx, clock = FakeEffects(), FakeClock()
    eng = Engine(fx, lambda: focus, clock=clock)
    return eng, fx, clock


def frame(v0=512, v1=512, v2=512, v3=512, v4=512):
    return [v0, v1, v2, v3, v4]


def test_raw_to_level_ley_cuadratica_y_mute():
    assert raw_to_level(0) == 0.0
    assert raw_to_level(MUTE_THRESHOLD) == 0.0
    assert raw_to_level(1023) == 1.0
    assert raw_to_level(511.5) == pytest.approx(0.25)


def test_linear_to_vu_y_pico():
    assert linear_to_vu(0) == 0.0
    assert linear_to_vu(1.0) == 1.0
    assert linear_to_vu(0.001) == 0.0          # -60 dB, por debajo del piso de -30
    assert next_peak(0.5, 0.2) == 0.5
    assert next_peak(0.1, 0.5) == pytest.approx(0.488)


def test_no_aplica_volumen_antes_del_hardware():
    eng, fx, _ = make()
    eng.set_assignments({0: "Discord"}, [])
    assert eng.tick() is False
    assert fx.calls == []


def test_deadband_ignora_ruido():
    eng, _, _ = make()
    eng.on_frame(frame(v0=512 + DEADBAND - 1))
    assert eng.target_vals[0] == 512.0
    eng.on_frame(frame(v0=512 + DEADBAND))
    assert eng.target_vals[0] == 512.0 + DEADBAND


def test_tick_converge_y_aplica_volumen():
    eng, fx, _ = make()
    eng.set_assignments({0: "Discord", 1: "master"}, [])
    eng.on_frame(frame(v0=800, v1=100))
    assert eng.tick() is True
    assert eng.smooth_vals[0] == 800.0
    assert ("set_volume", "Discord", lv(800)) in fx.calls
    assert ("set_volume", "master", lv(100)) in fx.calls


def test_canal_sin_asignar_no_hace_nada():
    eng, fx, _ = make()
    eng.on_frame(frame(v2=900))
    eng.tick()
    assert fx.volumes() == []


def test_throttle_40ms_y_valor_pendiente_se_aplica():
    eng, fx, clock = make()
    eng.set_assignments({0: "Spotify"}, [])
    eng.on_frame(frame(v0=600))
    eng.tick()
    clock.t += 0.010
    eng.on_frame(frame(v0=700))
    eng.tick()                                   # dentro de la ventana de 40 ms
    assert ("set_volume", "Spotify", lv(700)) not in fx.calls
    clock.t += 0.040
    assert eng.tick() is False                   # ya no se mueve...
    assert ("set_volume", "Spotify", lv(700)) in fx.calls   # ...pero el pendiente se aplica


def test_is_muted():
    eng, _, _ = make()
    eng.set_assignments({0: "Discord"}, [])
    eng.on_frame(frame(v0=10, v1=10))
    eng.tick()
    assert eng.is_muted(0) is True
    assert eng.is_muted(1) is False              # sin asignar nunca está "muteado"


def test_prioridad_ultimo_que_se_movio_gana():
    eng, fx, _ = make(focus="Discord")
    eng.set_assignments({0: "Discord", 4: "focus"}, [])
    eng.apply_channel(0, 500)
    assert fx.volumes()[-1] == ("set_volume", "Discord", lv(500))
    eng.apply_channel(4, 300)                    # el foco se mueve: toma el control
    assert fx.volumes()[-1] == ("set_volume", "Discord", lv(300))
    n = len(fx.volumes())
    eng.apply_channel(0, 520)                    # directo se mueve 20 (<=40): no recupera
    assert len(fx.volumes()) == n
    eng.apply_channel(0, 560)                    # se mueve 60 (>40): recupera
    assert fx.volumes()[-1] == ("set_volume", "Discord", lv(560))
    n = len(fx.volumes())
    eng.apply_channel(4, 305)                    # el foco apenas se mueve (<=15): no toma
    assert len(fx.volumes()) == n


def test_foco_sin_ventana_activa_no_aplica():
    eng, fx, _ = make(focus="")
    eng.set_assignments({4: "focus"}, [])
    eng.apply_channel(4, 700)
    assert fx.volumes() == []
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_engine.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.core.engine'`.

- [ ] **Paso 3: Implementar**

`knob/core/engine.py`:
```python
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
```

- [ ] **Paso 4: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_engine.py`
Esperado: `10 passed`.

- [ ] **Paso 5: Commit**

```bash
git add knob/core/engine.py tests/fakes.py tests/test_engine.py
git commit -m "feat: motor de suavizado, volumen y prioridad de foco portado de v3.3" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 3: Motor: botones de los encoders

**Archivos:**
- Modificar: `knob/core/engine.py` (agregar el import de `protocol` y los métodos `on_button` y `_execute`)
- Test: `tests/test_engine_buttons.py`

**Interfaces:**
- Consume: `Engine` de la Tarea 2, `protocol.BTN_CLICK/BTN_DBLCLICK/BTN_HELD_START/BTN_HELD_END`.
- Produce: `Engine.on_button(enc_idx: int, code: int) -> None`.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_engine_buttons.py`:
```python
# -*- coding: utf-8 -*-
from knob.core import protocol
from knob.core.engine import Engine, raw_to_level
from tests.fakes import FakeClock, FakeEffects


def lv(raw):
    return round(raw_to_level(raw), 4)


def make(channels, enc0):
    fx, clock = FakeEffects(), FakeClock()
    eng = Engine(fx, lambda: "", clock=clock)
    eng.set_assignments(channels, [enc0, {}])
    return eng, fx, clock


def settle(eng, clock, v3):
    """Manda un frame con el dial izquierdo en v3 y avanza un tick de 50 ms."""
    eng.on_frame([512, 512, 512, v3, 512])
    clock.t += 0.05
    eng.tick()


def test_click_play_pause_y_ninguna():
    eng, fx, _ = make({}, {"click": "play_pause", "dblclick": "none"})
    eng.on_button(0, protocol.BTN_CLICK)
    eng.on_button(0, protocol.BTN_DBLCLICK)
    assert fx.calls == [("media_key", "play_pause")]


def test_doble_click_atajo_de_teclado():
    eng, fx, _ = make({}, {"dblclick": "hotkey", "hotkey_keys_dblclick": "ctrl+shift+m"})
    eng.on_button(0, protocol.BTN_DBLCLICK)
    assert fx.calls == [("send_hotkey", "ctrl+shift+m")]


def test_atajo_vacio_no_hace_nada():
    eng, fx, _ = make({}, {"click": "hotkey", "hotkey_keys_click": ""})
    eng.on_button(0, protocol.BTN_CLICK)
    assert fx.calls == []


def test_vol_up_y_down():
    eng, fx, _ = make({}, {"click": "vol_up", "dblclick": "vol_down"})
    eng.on_button(0, protocol.BTN_CLICK)
    eng.on_button(0, protocol.BTN_DBLCLICK)
    assert fx.calls == [("step_master", 0.05), ("step_master", -0.05)]


def test_mute_guarda_y_restaura_e_ignora_hardware():
    eng, fx, clock = make({3: "master"}, {"click": "mute"})
    settle(eng, clock, 700)
    eng.on_button(0, protocol.BTN_CLICK)              # mute
    assert eng.dial_muted[0] is True
    assert eng.smooth_vals[3] == 0.0
    assert fx.calls[-1] == ("set_volume", "master", 0.0)
    settle(eng, clock, 800)                            # mover el encoder muteado no hace nada
    assert eng.target_vals[3] == 0.0
    assert eng.is_muted(3) is True
    eng.on_button(0, protocol.BTN_CLICK)              # unmute
    assert eng.dial_muted[0] is False
    assert eng.smooth_vals[3] == 700.0
    assert fx.calls[-1] == ("set_volume", "master", lv(700))


def test_mute_sin_asignacion_no_hace_nada():
    eng, fx, _ = make({}, {"click": "mute"})
    eng.on_button(0, protocol.BTN_CLICK)
    assert eng.dial_muted[0] is False
    assert fx.calls == []


def test_mantener_y_girar_comportamiento_v33():
    """v3.3: mientras se mantiene, el dial se congela y el destino de click_turn recibe
    el valor CONGELADO (no el giro). Al soltar se restaura y el próximo frame salta.
    Se conserva tal cual a propósito; ver la nota al final del plan."""
    eng, fx, clock = make({3: "master"}, {"click_turn": "mic"})
    settle(eng, clock, 600)
    eng.on_button(0, protocol.BTN_HELD_START)
    assert eng.hold_saved[0] == 600.0
    fx.calls.clear()
    settle(eng, clock, 900)
    assert eng.smooth_vals[3] == 600.0
    assert ("set_volume", "mic", lv(600)) in fx.calls
    assert ("set_volume", "master", lv(900)) not in fx.calls
    eng.on_button(0, protocol.BTN_HELD_END)
    assert eng.smooth_vals[3] == 600.0 and eng.target_vals[3] == 600.0
    assert eng.hold_saved[0] is None
    settle(eng, clock, 900)
    assert eng.smooth_vals[3] == 900.0


def test_mantener_sin_accion_mueve_normal():
    eng, fx, clock = make({3: "master"}, {"click_turn": "none"})
    settle(eng, clock, 600)
    eng.on_button(0, protocol.BTN_HELD_START)
    settle(eng, clock, 900)
    assert eng.smooth_vals[3] == 900.0
    assert ("set_volume", "master", lv(900)) in fx.calls


def test_mantener_y_girar_cambia_salida_con_antirrebote():
    eng, fx, clock = make({3: "master"}, {"click_turn": "output_device"})
    settle(eng, clock, 500)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    settle(eng, clock, 600)            # primera lectura: solo guarda referencia
    settle(eng, clock, 700)            # +100 (>35): siguiente dispositivo
    settle(eng, clock, 800)            # +100 pero dentro de 800 ms: ignorado
    clock.t += 1.0
    settle(eng, clock, 600)            # -200: dispositivo anterior
    assert [c for c in fx.calls if c[0] == "cycle_output"] == [("cycle_output", 1),
                                                                ("cycle_output", -1)]
    assert not any(c[0] == "set_volume" and c[1] == "master" for c in fx.calls)
    eng.on_button(0, protocol.BTN_HELD_END)
    assert eng.smooth_vals[3] == 500.0
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_engine_buttons.py`
Esperado: FAIL con `AttributeError: 'Engine' object has no attribute 'on_button'`.

- [ ] **Paso 3: Implementar**

En `knob/core/engine.py`, agregar debajo de `from typing import Callable, Protocol`:
```python
from knob.core import protocol
```

Agregar estos métodos al final de la clase `Engine`:
```python
    # --- botones de encoders --------------------------------------------
    def on_button(self, enc_idx: int, code: int) -> None:
        """code: 1=click, 2=doble click, 3=empieza mantener, 4=termina mantener."""
        ch = 3 + enc_idx
        if code == protocol.BTN_CLICK:
            self._execute(enc_idx, "click")
        elif code == protocol.BTN_DBLCLICK:
            self._execute(enc_idx, "dblclick")
        elif code == protocol.BTN_HELD_START:
            self.encoder_held[enc_idx] = True
            if self.turn_key(enc_idx) != "none":
                self.hold_saved[enc_idx] = self.smooth_vals[ch]
        elif code == protocol.BTN_HELD_END:
            self.encoder_held[enc_idx] = False
            if self.turn_key(enc_idx) != "none" and self.hold_saved[enc_idx] is not None:
                self.smooth_vals[ch] = self.hold_saved[enc_idx]
                self.target_vals[ch] = self.hold_saved[enc_idx]
            self.hold_saved[enc_idx] = None
            self.enc_prev_raw[enc_idx] = None

    def _execute(self, enc_idx: int, action_type: str) -> None:
        entry = self.encoder_assignments[enc_idx]
        action = entry.get(action_type, "none")
        if action == "none":
            return
        ch = 3 + enc_idx
        if action == "mute":
            assignment = self.channel_assignments.get(ch)
            if not assignment:
                return
            if not self.dial_muted[enc_idx]:
                self.dial_saved_val[enc_idx] = max(self.smooth_vals[ch], 16.0)
                self.smooth_vals[ch] = 0.0
                self.target_vals[ch] = 0.0
                self.dial_muted[enc_idx] = True
                self.apply_key(assignment, 0)
            else:
                self.dial_muted[enc_idx] = False
                saved = self.dial_saved_val[enc_idx]
                self.smooth_vals[ch] = saved
                self.target_vals[ch] = saved
                self.apply_key(assignment, saved)
        elif action in ("play_pause", "next_track", "prev_track"):
            self.effects.media_key(action)
        elif action == "vol_up":
            self.effects.step_master(MASTER_STEP)
        elif action == "vol_down":
            self.effects.step_master(-MASTER_STEP)
        elif action == "hotkey":
            combo = entry.get(f"hotkey_keys_{action_type}", "")
            if combo:
                self.effects.send_hotkey(combo)
```

- [ ] **Paso 4: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_engine.py tests/test_engine_buttons.py`
Esperado: `19 passed`.

- [ ] **Paso 5: Commit**

```bash
git add knob/core/engine.py tests/test_engine_buttons.py
git commit -m "feat: acciones de botones de encoders en el motor (mute, mantener, multimedia, atajos)" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 4: Configuración y perfiles

**Archivos:**
- Crear: `knob/core/paths.py`, `knob/core/log.py`, `knob/core/config.py`
- Test: `tests/test_config.py`

**Interfaces:**
- Consume: `engine.default_encoder_entry()`.
- Produce:
  - `paths.is_frozen() -> bool`, `paths.project_root() -> str`, `paths.app_dir() -> str`, `paths.resource_path(rel: str) -> str`.
  - `log.log(msg: str) -> None`, `log.LOG_PATH`.
  - `config.VALID_THEMES`, `config.DEFAULT_THEME = "Nebula"`, `config.default_config_path() -> str`, `ProfileError(ValueError)`.
  - `Config(path: str | None = None)` con:
    - Datos: `data: dict`, propiedad `active_profile: str`.
    - Lectura: `profile_names() -> list[str]`, `channel_assignments() -> dict[int, str|None]`, `encoder_assignments() -> list[dict]`.
    - Escritura: `set_channel_assignment(ch: int, value: str|None)`, `set_encoder_assignments(encoders: list[dict])`.
    - Perfiles: `switch_profile(name)`, `create_profile(name, channels: dict[int, str|None], encoders: list[dict]) -> str`, `rename_profile(old, new) -> str`, `delete_profile(name) -> bool`.
    - Propiedades con setter que guardan solas: `theme`, `minimize_to_tray: bool`, `last_port: str|None`.
    - Rutas de ejecutables: `exe_path(app_name) -> str|None`, `remember_exe_path(app_name, path)`.
    - `save()`.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_config.py`:
```python
# -*- coding: utf-8 -*-
import json

import pytest

from knob.core.config import Config, ProfileError


def write(path, data):
    path.write_text(json.dumps(data), encoding="utf-8")


def test_archivo_inexistente_usa_defaults_sin_crear_archivo(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    assert c.profile_names() == ["Default"]
    assert c.active_profile == "Default"
    assert c.theme == "Nebula"
    assert c.channel_assignments() == {0: None, 1: None, 2: None, 3: None, 4: None}
    assert c.encoder_assignments()[0]["click"] == "none"
    assert c.minimize_to_tray is False
    assert not p.exists()


def test_migra_formato_plano_a_perfiles(tmp_path):
    p = tmp_path / "config.json"
    write(p, {"channel_assignments": {"0": "Discord", "3": "master"}, "theme": "Cobalt"})
    c = Config(str(p))
    assert c.profile_names() == ["Default"]
    assert c.channel_assignments()[0] == "Discord"
    assert c.channel_assignments()[3] == "master"
    assert c.theme == "Cobalt"


@pytest.mark.parametrize("viejo,nuevo", [("Arctic", "Onyx"), ("Steel", "Onyx"),
                                          ("Raro", "Nebula"), ("Magma", "Magma")])
def test_migracion_de_temas(tmp_path, viejo, nuevo):
    p = tmp_path / "config.json"
    write(p, {"theme": viejo})
    assert Config(str(p)).theme == nuevo


def test_config_corrupto_no_se_pisa(tmp_path):
    p = tmp_path / "config.json"
    p.write_text("{no es json", encoding="utf-8")
    c = Config(str(p))
    assert c.profile_names() == ["Default"]
    assert p.read_text(encoding="utf-8") == "{no es json"


def test_asignacion_persiste(tmp_path):
    p = tmp_path / "config.json"
    Config(str(p)).set_channel_assignment(1, "Spotify")
    c = Config(str(p))
    assert c.channel_assignments()[1] == "Spotify"
    assert json.loads(p.read_text(encoding="utf-8"))["profiles"]["Default"]["channel_assignments"]["1"] == "Spotify"


def test_encoders_se_normalizan(tmp_path):
    p = tmp_path / "config.json"
    write(p, {"profiles": {"Default": {"encoder_assignments": [{"click": "mute"}]}},
              "active_profile": "Default"})
    encs = Config(str(p)).encoder_assignments()
    assert len(encs) == 2
    assert encs[0]["click"] == "mute"
    assert encs[0]["hotkey_keys_click"] == ""
    assert encs[1]["click_turn"] == "none"


def test_crear_renombrar_y_borrar_perfiles(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    c.set_channel_assignment(0, "Opera")
    c.create_profile("Gaming", c.channel_assignments(), c.encoder_assignments())
    assert c.active_profile == "Gaming"
    assert c.channel_assignments()[0] == "Opera"
    c.set_channel_assignment(0, "Discord")
    c.switch_profile("Default")
    assert c.channel_assignments()[0] == "Opera"
    assert c.rename_profile("Gaming", "Juegos") == "Juegos"
    assert c.profile_names() == ["Default", "Juegos"]
    c.switch_profile("Juegos")
    assert c.channel_assignments()[0] == "Discord"
    assert c.delete_profile("Juegos") is True
    assert c.active_profile == "Default"
    assert Config(str(p)).profile_names() == ["Default"]


def test_renombrar_activo_actualiza_activo(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    c.rename_profile("Default", "Casa")
    assert c.active_profile == "Casa"


def test_no_se_borra_el_ultimo_perfil(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    assert c.delete_profile("Default") is False
    assert c.profile_names() == ["Default"]


def test_nombres_de_perfil_invalidos(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    with pytest.raises(ProfileError):
        c.create_profile("   ", {}, [])
    with pytest.raises(ProfileError):
        c.create_profile("Default", {}, [])
    c.create_profile("Otro", {}, [])
    with pytest.raises(ProfileError):
        c.rename_profile("Otro", "Default")
    with pytest.raises(ProfileError):
        c.switch_profile("No existe")


def test_preferencias_y_exe_paths(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    c.theme = "Toxic"
    c.minimize_to_tray = True
    c.last_port = "COM13"
    c.remember_exe_path("Discord", r"C:\Discord\Discord.exe")
    c2 = Config(str(p))
    assert (c2.theme, c2.minimize_to_tray, c2.last_port) == ("Toxic", True, "COM13")
    assert c2.exe_path("discord") == r"C:\Discord\Discord.exe"
    with pytest.raises(ValueError):
        c2.theme = "Arctic"
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_config.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.core.config'`.

- [ ] **Paso 3: Implementar paths y log**

`knob/core/paths.py`:
```python
# -*- coding: utf-8 -*-
"""Rutas que funcionan igual desde el código fuente y desde el build de PyInstaller."""
import os
import sys


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def app_dir() -> str:
    """Carpeta del KNOB.exe instalado (o la raíz del proyecto en desarrollo)."""
    return os.path.dirname(sys.executable) if is_frozen() else project_root()


def resource_path(rel: str) -> str:
    """Archivos empaquetados (icon.png, fonts/, version.txt)."""
    base = getattr(sys, "_MEIPASS", None) if is_frozen() else project_root()
    return os.path.join(base or app_dir(), rel)
```

`knob/core/log.py`:
```python
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
```

- [ ] **Paso 4: Implementar config**

`knob/core/config.py`:
```python
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
                "channel_assignments": d.get("channel_assignments", {}) or {},
                "encoder_assignments": d.get("encoder_assignments", [{}, {}]),
            }}
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
                json.dump(self.data, f, indent=2, ensure_ascii=False)
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
```

- [ ] **Paso 5: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_config.py`
Esperado: `14 passed`.

- [ ] **Paso 6: Commit**

```bash
git add knob/core/paths.py knob/core/log.py knob/core/config.py tests/test_config.py
git commit -m "feat: configuración y perfiles compatibles con config.json de v3.3" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 5: Integración con Windows y actualizador

**Archivos:**
- Crear: `knob/core/system.py`, `knob/core/updater.py`
- Test: `tests/test_system.py`

**Interfaces:**
- Consume: `paths.app_dir`, `paths.resource_path`, `log.log`.
- Produce `system`:
  - Instancia única: `acquire_single_instance() -> bool`, `show_already_running()`.
  - Inicio con Windows: `startup_command() -> str`, `is_startup_enabled() -> bool`, `set_startup_enabled(enabled: bool)`, `refresh_startup_path()`.
  - Teclado: `parse_hotkey(combo) -> tuple[list[str], list[str]]`, `send_media_key(action: str)`, `send_hotkey(combo: str)`.
  - Utilidades: `strip_exe(name) -> str`, `FocusTracker` (`.start()`, `.current: str`), `hex_to_colorref(hex) -> int`, `style_title_bar(hwnd: int, caption_hex: str)`, `set_app_user_model_id(app_id: str)`.
- Produce `updater`: `RELEASES_URL`, `parse_version(v) -> tuple`, `read_version() -> str`, `check_for_update(current: str) -> tuple[str, str] | None`, `download(url, on_progress: Callable[[int], None]) -> str`, `launch_installer(path)`.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_system.py`:
```python
# -*- coding: utf-8 -*-
from knob.core import system, updater


def test_parse_hotkey():
    assert system.parse_hotkey("ctrl+shift+m") == (["ctrl", "shift"], ["m"])
    assert system.parse_hotkey(" Alt + F4 ") == (["alt"], ["f4"])
    assert system.parse_hotkey("") == ([], [])


def test_strip_exe():
    assert system.strip_exe("Discord.exe") == "Discord"
    assert system.strip_exe("spotify.EXE") == "spotify"
    assert system.strip_exe("opera") == "opera"
    assert system.strip_exe("") == ""


def test_hex_to_colorref():
    assert system.hex_to_colorref("#0C0A14") == 0x00140A0C
    assert system.hex_to_colorref("FF0000") == 0x000000FF


def test_parse_version():
    assert updater.parse_version("v4.0.0") == (4, 0, 0)
    assert updater.parse_version("v3.10.0") > updater.parse_version("v3.9.9")
    assert updater.parse_version("basura") == (0,)


def test_read_version_lee_version_txt():
    assert updater.read_version().startswith("v")
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_system.py`
Esperado: FAIL con `ImportError: cannot import name 'system'`.

- [ ] **Paso 3: Implementar system**

`knob/core/system.py`:
```python
# -*- coding: utf-8 -*-
"""Integración con Windows: instancia única, inicio automático, teclas, app en foco y barra de título."""
import ctypes
import os
import sys
import threading
import time
import winreg

import psutil

from knob.core.log import log

# --- instancia única ---------------------------------------------------------
_MUTEX_NAME = "Global\\KNOBAppMutex"
_ERROR_ALREADY_EXISTS = 183
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_mutex_handle = None


def acquire_single_instance() -> bool:
    global _mutex_handle
    _mutex_handle = _kernel32.CreateMutexW(None, False, _MUTEX_NAME)
    return ctypes.get_last_error() != _ERROR_ALREADY_EXISTS


def show_already_running() -> None:
    ctypes.windll.user32.MessageBoxW(0, "KNOB ya está en ejecución.", "KNOB", 0x30)


def set_app_user_model_id(app_id: str) -> None:
    """Hace que la barra de tareas muestre el ícono de KNOB y no el de Python."""
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


# --- inicio con Windows ------------------------------------------------------
APP_REG_NAME = "KNOBApp"
_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def startup_command() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --minimized'
    return f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}" --minimized'


def is_startup_enabled() -> bool:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, APP_REG_NAME)
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def set_startup_enabled(enabled: bool) -> None:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_SET_VALUE)
        if enabled:
            winreg.SetValueEx(key, APP_REG_NAME, 0, winreg.REG_SZ, startup_command())
        else:
            try:
                winreg.DeleteValue(key, APP_REG_NAME)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception as e:
        log(f"inicio con Windows: {e}")


def refresh_startup_path() -> None:
    """Si el inicio automático está activo, lo apunta al ejecutable actual."""
    if is_startup_enabled():
        set_startup_enabled(True)


# --- teclado -----------------------------------------------------------------
_KEYEVENTF_EXTENDEDKEY = 0x0001
_KEYEVENTF_KEYUP = 0x0002
_MODIFIERS = ("ctrl", "shift", "alt", "win")
VK_MAP = {
    "ctrl": 0x11, "shift": 0x10, "alt": 0x12, "win": 0x5B,
    "space": 0x20, "enter": 0x0D, "return": 0x0D, "esc": 0x1B, "escape": 0x1B,
    "tab": 0x09, "backspace": 0x08, "delete": 0x2E, "insert": 0x2D,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "prior": 0x21,
    "pagedown": 0x22, "next": 0x22,
    "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
    "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
}
MEDIA_VK = {"play_pause": 0xB3, "next_track": 0xB0, "prev_track": 0xB1}


def parse_hotkey(combo: str):
    parts = [p.strip().lower() for p in (combo or "").split("+") if p.strip()]
    return [p for p in parts if p in _MODIFIERS], [p for p in parts if p not in _MODIFIERS]


def send_media_key(action: str) -> None:
    vk = MEDIA_VK.get(action)
    if vk is None:
        return
    ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY, 0)
    ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP, 0)


def _vk_for(key: str):
    if key in VK_MAP:
        return VK_MAP[key]
    if len(key) == 1:
        return ctypes.windll.user32.VkKeyScanW(ord(key)) & 0xFF
    return None


def send_hotkey(combo: str) -> None:
    modifiers, keys = parse_hotkey(combo)
    if not keys:
        return
    user32 = ctypes.windll.user32
    try:
        for mod in modifiers:
            user32.keybd_event(VK_MAP[mod], 0, 0, 0)
        for k in keys:
            vk = _vk_for(k)
            if vk:
                user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY, 0)
                user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP, 0)
        for mod in reversed(modifiers):
            user32.keybd_event(VK_MAP[mod], 0, _KEYEVENTF_KEYUP, 0)
    except Exception as e:
        log(f"send_hotkey({combo}): {e}")


# --- app en foco -------------------------------------------------------------
def strip_exe(name: str) -> str:
    return name[:-4] if name and name.lower().endswith(".exe") else (name or "")


class FocusTracker:
    """Hilo que cada 150 ms guarda el último exe en primer plano que no sea KNOB."""

    def __init__(self, interval_s: float = 0.15):
        self.current = ""
        self._interval = interval_s

    def start(self) -> None:
        threading.Thread(target=self._run, daemon=True, name="focus-tracker").start()

    def _run(self) -> None:
        own_pid = os.getpid()
        user32 = ctypes.windll.user32
        while True:
            try:
                hwnd = user32.GetForegroundWindow()
                if hwnd:
                    pid = ctypes.c_ulong()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value != own_pid:
                        exe = strip_exe(psutil.Process(pid.value).name()).capitalize()
                        if exe:
                            self.current = exe
            except Exception:
                pass
            time.sleep(self._interval)


# --- barra de título ---------------------------------------------------------
_DWMWA_USE_IMMERSIVE_DARK_MODE = 20
_DWMWA_BORDER_COLOR = 34
_DWMWA_CAPTION_COLOR = 35


def hex_to_colorref(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


def style_title_bar(hwnd: int, caption_hex: str) -> None:
    """Barra de título oscura y del color del tema (el color solo aplica en Windows 11)."""
    try:
        dwm = ctypes.windll.dwmapi
        dark = ctypes.c_int(1)
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(dark), ctypes.sizeof(dark))
        color = ctypes.c_uint(hex_to_colorref(caption_hex))
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_CAPTION_COLOR, ctypes.byref(color), ctypes.sizeof(color))
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_BORDER_COLOR, ctypes.byref(color), ctypes.sizeof(color))
    except Exception as e:
        log(f"style_title_bar: {e}")
```

- [ ] **Paso 4: Implementar updater**

`knob/core/updater.py`:
```python
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
```

- [ ] **Paso 5: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_system.py`
Esperado: `5 passed`.

- [ ] **Paso 6: Verificación manual rápida (Windows real)**

Correr:
```bash
venv/Scripts/python.exe -c "from knob.core import system, updater; print(system.is_startup_enabled(), updater.read_version())"
```
Esperado: imprime `True v3.3.0` o `False v3.3.0` sin errores.

- [ ] **Paso 7: Commit**

```bash
git add knob/core/system.py knob/core/updater.py tests/test_system.py
git commit -m "feat: integración con Windows (instancia única, inicio, teclas, foco, barra de título) y actualizador" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 6: Audio de Windows (pycaw)

**Archivos:**
- Crear: `knob/core/audio.py`
- Test: `tests/test_audio.py` (solo la parte pura) + verificación manual en Windows

**Interfaces:**
- Consume: `log.log`, `system.strip_exe`.
- Produce:
  - Funciones: `is_system_app(name) -> bool`, `find_exe_path(app_name) -> str|None`, `set_default_output_device(device_id)`.
  - `AudioBackend` con:
    - Ciclo de vida: `start()`, `refresh()`, `invalidate()`.
    - Consulta: `app_names() -> list[str]` (refresca en el momento), `master_level() -> float`, `peak(assignment: str|None, focus_exe: str) -> float`.
    - Volumen: `set_volume(target, level)`, `set_master(level)`, `set_mic(level)`, `set_app(name, level)`, `step_master(delta)`.
  - `OutputDevices` con `load() -> str|None` (nombre del dispositivo predeterminado; correrlo en un hilo), `cycle(direction) -> str|None` y `current_name() -> str|None`.

Los nombres de las apps son el nombre del exe sin `.exe` y con `capitalize()` (ej. `Discord`), igual que v3.3, para que las asignaciones guardadas sigan funcionando.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_audio.py`:
```python
# -*- coding: utf-8 -*-
from knob.core.audio import OutputDevices, is_system_app


def test_lista_negra_de_procesos_del_sistema():
    assert is_system_app("audiodg")
    assert is_system_app("MsEdgeWebView2")
    assert not is_system_app("Discord")


def test_cycle_sin_dispositivos_devuelve_none():
    assert OutputDevices().cycle(1) is None


def test_cycle_rota_indices(monkeypatch):
    od = OutputDevices()
    od.devices = [("id1", "Parlantes"), ("id2", "Auriculares")]
    od.index = 0
    monkeypatch.setattr(OutputDevices, "_set_default_async", staticmethod(lambda dev_id: None))
    assert od.cycle(1) == "Auriculares"
    assert od.cycle(1) == "Parlantes"
    assert od.cycle(-1) == "Auriculares"
    assert od.current_name() == "Auriculares"
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_audio.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.core.audio'`.

- [ ] **Paso 3: Implementar**

`knob/core/audio.py`:
```python
# -*- coding: utf-8 -*-
"""Audio de Windows vía pycaw: volumen por app, master, micrófono, medidores y salida predeterminada.
Portado de app.py v3.3 (mismas llamadas a pycaw/COM)."""
import ctypes
import threading
import winreg

import comtypes
import psutil
from comtypes import CLSCTX_ALL
from pycaw.pycaw import (AudioUtilities, IAudioEndpointVolume, IAudioMeterInformation,
                         ISimpleAudioVolume)

from knob.core.log import log
from knob.core.system import strip_exe

CACHE_TTL_S = 1.5

# Procesos del sistema que nunca deben aparecer en la lista de apps con audio
SYSTEM_APP_BLACKLIST = {
    "msedgewebview2", "amdrsserv", "amdow", "amdnotificationapp",
    "audiodg", "svchost", "csrss", "dwm", "rundll32", "conhost",
    "searchhost", "shellexperiencehost", "startmenuexperiencehost",
    "runtimebroker", "applicationframehost", "wudfhost",
    "sihost", "fontdrvhost", "wmiprvse", "wlanext",
    "nvcontainer", "nvdisplay.container", "nvsphelper64",
    "igfxem", "igfxext", "igfxhk", "igfxtray",
    "asusdec", "asusoptimization", "asusupdateexe",
    "gamebar", "gamebarftserver", "gamingservices",
    "microsoftedge", "msedge",
    "widgets", "widgetsservice",
}


def is_system_app(name: str) -> bool:
    return name.lower() in SYSTEM_APP_BLACKLIST


def find_exe_path(app_name: str):
    """Ruta del exe de una app que está corriendo (para extraer su ícono)."""
    target = app_name.lower()
    try:
        for proc in psutil.process_iter(["name", "exe"]):
            try:
                if strip_exe(proc.info.get("name") or "").lower() == target and proc.info.get("exe"):
                    return proc.info["exe"]
            except Exception:
                continue
    except Exception:
        pass
    return None


class AudioBackend:
    def __init__(self):
        self._sessions = {}       # app -> [ISimpleAudioVolume, ...] (Discord tiene varias sesiones)
        self._meters = {}         # app -> [IAudioMeterInformation, ...]
        self._master_meter = None
        self._mic_meter = None
        self._wake = threading.Event()

    # --- caché de sesiones -------------------------------------------------
    def start(self) -> None:
        threading.Thread(target=self._refresh_loop, daemon=True, name="audio-cache").start()

    def _refresh_loop(self) -> None:
        comtypes.CoInitialize()
        while True:
            try:
                self.refresh()
            except Exception as e:
                log(f"audio: error refrescando sesiones: {e}")
            self._wake.wait(CACHE_TTL_S)
            self._wake.clear()

    def invalidate(self) -> None:
        """Una sesión murió: adelanta el próximo refresco."""
        self._wake.set()

    def refresh(self) -> None:
        new_vol, new_meter = {}, {}
        for s in AudioUtilities.GetAllSessions():
            if not s.Process:
                continue
            raw = strip_exe(s.Process.name())
            if not raw or is_system_app(raw):
                continue
            name = raw.capitalize()
            new_vol.setdefault(name, [])
            new_meter.setdefault(name, [])
            try:
                new_vol[name].append(s._ctl.QueryInterface(ISimpleAudioVolume))
            except Exception:
                pass
            try:
                new_meter[name].append(s._ctl.QueryInterface(IAudioMeterInformation))
            except Exception:
                pass
        self._sessions = new_vol
        self._meters = new_meter
        try:
            spk = AudioUtilities.GetSpeakers()
            self._master_meter = spk.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None
                                              ).QueryInterface(IAudioMeterInformation)
        except Exception:
            pass
        try:
            mic = AudioUtilities.GetMicrophone()
            if mic:
                self._mic_meter = mic.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None
                                               ).QueryInterface(IAudioMeterInformation)
        except Exception:
            pass

    def app_names(self) -> list:
        try:
            self.refresh()
        except Exception as e:
            log(f"audio: app_names: {e}")
        return list(self._sessions.keys())

    # --- volumen -----------------------------------------------------------
    def set_volume(self, target: str, level: float) -> None:
        if target == "master":
            self.set_master(level)
        elif target == "mic":
            self.set_mic(level)
        else:
            self.set_app(target, level)

    def set_master(self, level: float) -> None:
        AudioUtilities.GetSpeakers().EndpointVolume.SetMasterVolumeLevelScalar(level, None)

    def master_level(self) -> float:
        return AudioUtilities.GetSpeakers().EndpointVolume.GetMasterVolumeLevelScalar()

    def step_master(self, delta: float) -> None:
        self.set_master(max(0.0, min(1.0, self.master_level() + delta)))

    def set_mic(self, level: float) -> None:
        mic = AudioUtilities.GetMicrophone()
        if mic:
            interface = mic.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            interface.QueryInterface(IAudioEndpointVolume).SetMasterVolumeLevelScalar(level, None)

    def set_app(self, name: str, level: float) -> None:
        failed = False
        for ctl in self._sessions.get(name, []):
            try:
                ctl.SetMasterVolume(level, None)
            except Exception:
                failed = True
        if failed:
            self.invalidate()

    # --- medidores ---------------------------------------------------------
    def peak(self, assignment, focus_exe: str) -> float:
        """Pico de audio (0-1) del destino de un canal."""
        if not assignment:
            return 0.0
        try:
            if assignment == "master":
                return self._master_meter.GetPeakValue() if self._master_meter else 0.0
            if assignment == "mic":
                return self._mic_meter.GetPeakValue() if self._mic_meter else 0.0
            if assignment == "focus":
                wanted = (focus_exe or "").lower()
                key = next((k for k in self._meters if k.lower() == wanted), None)
                meters = self._meters.get(key, []) if key else []
            else:
                meters = self._meters.get(assignment, [])
            peak, failed = 0.0, False
            for m in meters:
                try:
                    peak = max(peak, m.GetPeakValue())
                except Exception:
                    failed = True
            if failed:
                self.invalidate()
            return peak
        except Exception:
            return 0.0


# --- dispositivos de salida ----------------------------------------------------
def _device_friendly_name(device_id: str):
    """Nombre completo tipo 'Speakers (UG-08S)' leído del registro."""
    try:
        guid = device_id.rsplit(".", 1)[-1]
        key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{guid}\Properties"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | 0x0100)
        endpoint, _ = winreg.QueryValueEx(key, "{a45c254e-df1c-4efd-8020-67d146a850e0},2")
        try:
            device, _ = winreg.QueryValueEx(key, "{b3f8fa53-0004-438e-9003-51a46e139bfc},6")
            name = f"{endpoint} ({device})"
        except Exception:
            name = endpoint
        winreg.CloseKey(key)
        return name
    except Exception:
        return None


def set_default_output_device(device_id: str) -> None:
    """Cambia la salida predeterminada con la interfaz COM no documentada IPolicyConfig."""
    try:
        from comtypes import GUID
        clsid = GUID("{870AF99C-171D-4F9E-AF0D-E63DF40C2BC9}")
        iid_unk = GUID("{00000000-0000-0000-C000-000000000046}")
        iid_policy = GUID("{F8679F50-850A-41CF-9C72-430F290290C8}")
        unk_ptr = ctypes.POINTER(ctypes.c_void_p)()
        hr = ctypes.windll.ole32.CoCreateInstance(
            ctypes.byref(clsid), None, 1, ctypes.byref(iid_unk), ctypes.byref(unk_ptr))
        if hr != 0:
            return
        vtable_unk = ctypes.cast(unk_ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
        query_interface = ctypes.WINFUNCTYPE(
            ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p))(vtable_unk[0])
        policy_ptr = ctypes.c_void_p()
        hr = query_interface(ctypes.cast(unk_ptr, ctypes.c_void_p).value,
                             ctypes.addressof(iid_policy), ctypes.byref(policy_ptr))
        if hr != 0 or not policy_ptr:
            return
        vtable = ctypes.cast(policy_ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
        set_default_endpoint = ctypes.WINFUNCTYPE(
            ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint)(vtable[13])
        for role in range(3):
            set_default_endpoint(policy_ptr.value, device_id, role)
        ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtable[2])(policy_ptr.value)
    except Exception as e:
        log(f"set_default_output_device: {e}")


class OutputDevices:
    def __init__(self):
        self.devices = []        # [(device_id, nombre), ...]
        self.index = 0
        self._lock = threading.Lock()

    def load(self):
        """Enumera las salidas activas. Correr en un hilo. Devuelve el nombre de la predeterminada."""
        comtypes.CoInitialize()
        try:
            enumerator = AudioUtilities.GetDeviceEnumerator()
            collection = enumerator.EnumAudioEndpoints(0, 1)   # eRender, ACTIVE
            devs = []
            for i in range(collection.GetCount()):
                dev_id = collection.Item(i).GetId()
                devs.append((dev_id, _device_friendly_name(dev_id) or f"Dispositivo {i + 1}"))
            idx = 0
            try:
                default_id = enumerator.GetDefaultAudioEndpoint(0, 1).GetId()
                idx = next((i for i, (d, _) in enumerate(devs) if d == default_id), 0)
            except Exception:
                pass
            with self._lock:
                self.devices, self.index = devs, idx
            return self.current_name()
        except Exception as e:
            log(f"OutputDevices.load: {e}")
            return None

    def current_name(self):
        with self._lock:
            return self.devices[self.index][1] if self.devices else None

    def cycle(self, direction: int):
        with self._lock:
            if not self.devices:
                return None
            self.index = (self.index + direction) % len(self.devices)
            dev_id, name = self.devices[self.index]
        self._set_default_async(dev_id)
        return name

    @staticmethod
    def _set_default_async(dev_id: str) -> None:
        def run():
            comtypes.CoInitialize()
            set_default_output_device(dev_id)
        threading.Thread(target=run, daemon=True).start()


if __name__ == "__main__":
    # Verificación manual: python -m knob.core.audio
    backend = AudioBackend()
    print("Apps con audio:", backend.app_names())
    print("Volumen master:", round(backend.master_level(), 3))
    od = OutputDevices()
    print("Salida actual:", od.load())
    print("Salidas:", [n for _, n in od.devices])
```

- [ ] **Paso 4: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_audio.py`
Esperado: `3 passed`.

- [ ] **Paso 5: Verificación manual en Windows**

Con Spotify o Discord reproduciendo algo, correr: `venv/Scripts/python.exe -m knob.core.audio`
Esperado: lista con esas apps (ej. `['Spotify', 'Discord']`), un volumen master entre 0 y 1, y el nombre de tu salida (ej. `Realtek Digital Output (Realtek(R) Audio)`). Sin excepciones.

- [ ] **Paso 6: Commit**

```bash
git add knob/core/audio.py tests/test_audio.py
git commit -m "feat: backend de audio pycaw y salidas de audio portados de v3.3" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 7: Conexión serial (búsqueda, handshake y lectura)

**Archivos:**
- Crear: `knob/core/serial_link.py`
- Test: `tests/test_serial_link.py`

**Interfaces:**
- Consume: `protocol.*`, `log.log`.
- Produce:
  - `handshake(ser, timeout_s=12.0, clock=time.time) -> bool`.
  - `candidate_ports(ports, last_port) -> list[str]`.
  - `SerialLink(last_port_getter: Callable[[], str|None])` (QObject) con:
    - Señales `button_event(int, int)`, `status_changed(str, str)` (estado `"searching"`, `"connected"` o `"busy"`, detalle = puerto) y `port_connected(str)`.
    - Métodos `start()`, `stop()` y `take_latest() -> list[int] | None` (último frame sin leer; `None` si no hay nuevo).

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_serial_link.py`:
```python
# -*- coding: utf-8 -*-
from types import SimpleNamespace

from knob.core import protocol
from knob.core.serial_link import candidate_ports, handshake


class FakeSerial:
    def __init__(self, lines):
        self._lines = list(lines)
        self.writes = []

    def readline(self):
        return (self._lines.pop(0) + "\n").encode() if self._lines else b""

    def write(self, data):
        self.writes.append(data)


class StepClock:
    """Avanza 0.3 s en cada consulta, para simular el paso del tiempo sin esperar."""
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        self.t += 0.3
        return self.t


def test_handshake_completo():
    ser = FakeSerial([protocol.CMD_AWAIT_INIT, protocol.LOG_INITIALIZED])
    assert handshake(ser, clock=StepClock()) is True
    assert ser.writes == [protocol.CMD_START]


def test_handshake_con_arduino_ya_transmitiendo():
    ser = FakeSerial(["512|512|512|512|512|0|0"])
    assert handshake(ser, clock=StepClock()) is True
    assert ser.writes == []


def test_handshake_silencio_manda_sondas_y_falla():
    ser = FakeSerial([])
    assert handshake(ser, timeout_s=5.0, clock=StepClock()) is False
    assert len(ser.writes) >= 3
    assert set(ser.writes) == {protocol.CMD_START}


def port(device, vid=None, description=""):
    return SimpleNamespace(device=device, vid=vid, description=description)


def test_candidate_ports_filtra_y_prioriza_el_ultimo():
    ports = [port("COM1", None, "Communications Port"),
             port("COM5", 0x2341, "Arduino Micro"),
             port("COM7", None, "USB-SERIAL CH340"),
             port("COM13", 0x1B4F, "SparkFun Pro Micro")]
    assert candidate_ports(ports, "COM13") == ["COM13", "COM5", "COM7"]
    assert candidate_ports(ports, None) == ["COM5", "COM7", "COM13"]
    assert candidate_ports([port("COM1")], None) == []
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_serial_link.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.core.serial_link'`.

- [ ] **Paso 3: Implementar**

`knob/core/serial_link.py`:
```python
# -*- coding: utf-8 -*-
"""Conexión con el KNOB: busca el puerto, hace el handshake y lee frames en un hilo.

El hilo nunca toca la interfaz. Los eventos salen por señales Qt (que llegan encoladas
al hilo principal) y el último frame se toma con take_latest() desde el timer de 16 ms.
"""
import threading
import time

import serial
import serial.tools.list_ports
from PySide6.QtCore import QObject, Signal

from knob.core import protocol
from knob.core.log import log

USB_KEYWORDS = ("USB", "CH340", "CP210", "FTDI", "ARDUINO")
HANDSHAKE_TIMEOUT_S = 12.0
PROBE_INTERVAL_S = 1.0
RETRY_DELAY_S = 2.0


def handshake(ser, timeout_s: float = HANDSHAKE_TIMEOUT_S, clock=time.time) -> bool:
    """Completa el handshake sobre un puerto ya abierto.
    Lee primero: si el Arduino ya está transmitiendo (reinicio de la app) conecta sin escribir."""
    t_end = clock() + timeout_s
    last_probe = clock()
    while clock() < t_end:
        line = ser.readline().decode(errors="ignore").strip()
        if line:
            if line == protocol.CMD_AWAIT_INIT:
                ser.write(protocol.CMD_START)
                last_probe = clock()
            elif line == protocol.LOG_INITIALIZED:
                return True
            elif protocol.is_stream_line(line):
                return True
        else:
            now = clock()
            if now - last_probe >= PROBE_INTERVAL_S:
                try:
                    ser.write(protocol.CMD_START)
                except serial.SerialTimeoutException:
                    log("handshake: write timeout")
                last_probe = now
    return False


def candidate_ports(ports, last_port) -> list:
    """Solo puertos USB/Arduino reales (nunca COM1 pelado); el último usado va primero."""
    cands = [p for p in ports
             if p.vid is not None
             or any(kw in (p.description or "").upper() for kw in USB_KEYWORDS)]
    cands.sort(key=lambda p: 0 if p.device == last_port else 1)
    return [p.device for p in cands]


class SerialLink(QObject):
    button_event = Signal(int, int)
    status_changed = Signal(str, str)
    port_connected = Signal(str)

    def __init__(self, last_port_getter, parent=None):
        super().__init__(parent)
        self._last_port = last_port_getter
        self._running = False
        self._ser = None
        self._latest = None
        self._latest_lock = threading.Lock()

    def start(self) -> None:
        if self._running:
            return
        self._running = True
        threading.Thread(target=self._run, daemon=True, name="serial-link").start()

    def stop(self) -> None:
        self._running = False
        ser = self._ser
        if ser is not None:
            try:
                ser.close()
            except Exception:
                pass

    def take_latest(self):
        with self._latest_lock:
            vals, self._latest = self._latest, None
        return vals

    # --- hilo -------------------------------------------------------------
    def _run(self) -> None:
        while self._running:
            try:
                ser = self._search()
                if ser is not None:
                    self._ser = ser
                    self._read_loop(ser)
            except Exception as e:
                log(f"serial: error inesperado: {e}")
                time.sleep(RETRY_DELAY_S)
            finally:
                if self._ser is not None:
                    try:
                        self._ser.close()
                    except Exception:
                        pass
                    self._ser = None
                    if self._running:
                        log("Conexión perdida, reiniciando búsqueda...")
                        self.status_changed.emit("searching", "")

    def _search(self):
        """Prueba los puertos candidatos una vez. Devuelve el Serial conectado o None."""
        ports = list(serial.tools.list_ports.comports())
        log(f"Puertos: {[(p.device, p.description, p.vid) for p in ports]}")
        cands = candidate_ports(ports, self._last_port())
        if not cands:
            log("No hay puertos Arduino/USB disponibles")
            self.status_changed.emit("searching", "")
            time.sleep(RETRY_DELAY_S)
            return None
        any_busy = False
        for dev in cands:
            if not self._running:
                return None
            log(f"Probando: {dev}")
            try:
                ser = serial.Serial(dev, protocol.BAUD_RATE, timeout=0.5, write_timeout=1.0)
            except Exception as e:
                busy = "Access is denied" in str(e) or "PermissionError" in str(e)
                log(f"{dev}: no se pudo abrir ({'OCUPADO' if busy else e})")
                if busy:
                    any_busy = True
                    self.status_changed.emit("busy", dev)
                continue
            if handshake(ser):
                log(f"Conectado en {dev}")
                self.status_changed.emit("connected", dev)
                self.port_connected.emit(dev)
                return ser
            log(f"{dev}: timeout/fallo, cerrando")
            try:
                ser.close()
            except Exception:
                pass
        if not any_busy:
            self.status_changed.emit("searching", "")
        log("KNOB no encontrado, reintentando en 2s...")
        time.sleep(RETRY_DELAY_S)
        return None

    def _read_loop(self, ser) -> None:
        ser.timeout = 0.05
        detector = protocol.ButtonEdgeDetector()
        while self._running and ser.is_open:
            try:
                raw = ser.readline()
            except Exception as e:
                log(f"serial: lectura falló: {e}")
                return
            line = raw.decode(errors="ignore").strip()
            if not line:
                continue
            frame = protocol.parse_frame(line)
            if frame is None:
                continue
            with self._latest_lock:
                self._latest = frame.values
            for enc_idx, code in detector.feed(frame.buttons):
                self.button_event.emit(enc_idx, code)
```

- [ ] **Paso 4: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_serial_link.py`
Esperado: `4 passed`.

- [ ] **Paso 5: Commit**

```bash
git add knob/core/serial_link.py tests/test_serial_link.py
git commit -m "feat: enlace serial con búsqueda de puerto, handshake y lectura en hilo" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 8: Tema y widgets base pintados a mano

**Archivos:**
- Crear: `knob/ui/theme.py`, `knob/ui/widgets/glass.py`, `fader.py`, `vu_meter.py`, `knob_dial.py`, `brand.py`, `status_pill.py`, `toggle.py`, `icon_badge.py` (todos en `knob/ui/widgets/` salvo theme)
- Test: `tests/test_widgets.py`

**Interfaces:**
- Consume: `paths.resource_path`.
- Produce `theme`:
  - Constantes: `TEXT`, `GREEN`, `AMBER`, `RED`, `GREEN_TEXT`, `AMBER_TEXT`, `RED_TEXT`.
  - `Theme` (dataclass congelado: `name`, `grad_a`, `grad_b`, `halo`, `bg`, `knob_face`, `knob_ring`, `pointer`, `brand_a`, `brand_b`, `text_dim`), `THEMES: dict[str, Theme]`, `THEME_ORDER: tuple`, `get_theme(name) -> Theme`.
  - Helpers: `qcolor(hex, alpha=255) -> QColor`, `rgba_css(hex, alpha: int) -> str`, `ui_font(px, weight=QFont.Weight.Normal) -> QFont`, `icon_font(px) -> QFont`, `brand_font(px) -> QFont`, `build_qss(theme) -> str`.
  - Glifos (Segoe Fluent Icons / MDL2): `GLYPH_VOLUME`, `GLYPH_MIC`, `GLYPH_FOCUS`, `GLYPH_ADD`, `GLYPH_SETTINGS`, `GLYPH_LEFT`, `GLYPH_RIGHT`, `GLYPH_CLOSE`.
- Produce widgets:
  - `GlassCard(parent=None, clickable=True)` con señal `clicked` y `set_dashed(bool)`. `Backdrop` con `set_theme(t)`.
  - `Fader(theme)` con `set_value(0..1)`, `set_muted(bool)`, `set_active(bool)` y `set_theme(t)`. `VuMeter()` con `set_level(level, peak)` y `set_active(bool)`.
  - `KnobDial(theme)` con `set_value(0..1)`, `set_muted(bool)`, `set_active(bool)` y `set_theme(t)`.
  - `BrandLabel(theme, px=26)` con `set_theme(t)`. `StatusPill()` con `set_state(state, text)` y `state()`.
  - `Toggle(theme)` (QAbstractButton chequeable) con `set_theme(t)`.
  - `IconBadge(size=28)` con `set_content(name: str, pixmap: QPixmap|None, kind: str|None)` (kind: `None`, `"master"`, `"mic"`, `"focus"` o `"unassigned"`), `set_dimmed(bool)` y `set_theme(t)`.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_widgets.py`:
```python
# -*- coding: utf-8 -*-
import pytest


def test_temas_completos(qapp):
    from knob.core.config import VALID_THEMES
    from knob.ui.theme import THEME_ORDER, THEMES, build_qss, get_theme
    assert set(THEMES) == set(VALID_THEMES) == set(THEME_ORDER)
    assert get_theme("Inexistente").name == "Nebula"
    for t in THEMES.values():
        assert "QMenu" in build_qss(t)


def test_rgba_css():
    from knob.ui.theme import rgba_css
    assert rgba_css("#7C3AED", 115) == "rgba(124,58,237,115)"


@pytest.mark.parametrize("theme_name", ["Nebula", "Onyx", "Sunset"])
def test_widgets_se_pintan(qapp, theme_name):
    from PySide6.QtGui import QColor, QPixmap
    from knob.ui.theme import get_theme
    from knob.ui.widgets.brand import BrandLabel
    from knob.ui.widgets.fader import Fader
    from knob.ui.widgets.glass import Backdrop, GlassCard
    from knob.ui.widgets.icon_badge import IconBadge
    from knob.ui.widgets.knob_dial import KnobDial
    from knob.ui.widgets.status_pill import StatusPill
    from knob.ui.widgets.toggle import Toggle
    from knob.ui.widgets.vu_meter import VuMeter

    t = get_theme(theme_name)
    fader = Fader(t)
    fader.resize(28, 200)
    for v, muted, active in [(0.0, False, True), (0.7, False, True), (0.5, True, True), (0.3, False, False)]:
        fader.set_value(v)
        fader.set_muted(muted)
        fader.set_active(active)
        assert not fader.grab().isNull()

    vu = VuMeter()
    vu.resize(4, 200)
    vu.set_level(0.9, 0.95)
    assert not vu.grab().isNull()

    dial = KnobDial(t)
    dial.set_value(0.75)
    dial.set_muted(True)
    assert not dial.grab().isNull()

    pill = StatusPill()
    for state, text in [("connected", "Conectado"), ("searching", "Buscando KNOB…"),
                        ("busy", "COM5 ocupado: cerrá el Serial Monitor")]:
        pill.set_state(state, text)
        assert pill.state() == state
        assert not pill.grab().isNull()

    toggle = Toggle(t)
    toggle.setChecked(True)
    assert not toggle.grab().isNull()

    badge = IconBadge(28)
    pm = QPixmap(64, 64)
    pm.fill(QColor("#1DB954"))
    for name, pix, kind in [("Spotify", pm, None), ("Discord", None, None),
                            ("", None, "master"), ("", None, "mic"), ("", None, "focus"),
                            ("", None, "unassigned")]:
        badge.set_content(name, pix, kind)
        badge.set_theme(t)
        assert not badge.grab().isNull()

    card = GlassCard()
    card.resize(150, 300)
    card.set_dashed(True)
    assert not card.grab().isNull()

    back = Backdrop(t)
    back.resize(520, 620)
    assert not back.grab().isNull()
    assert not BrandLabel(t).grab().isNull()


def test_glass_card_emite_clicked(qapp):
    from PySide6.QtCore import QPointF, Qt
    from PySide6.QtGui import QMouseEvent
    from PySide6.QtWidgets import QApplication
    from knob.ui.widgets.glass import GlassCard

    card = GlassCard()
    card.resize(100, 100)
    hits = []
    card.clicked.connect(lambda: hits.append(1))
    for etype in (QMouseEvent.Type.MouseButtonPress, QMouseEvent.Type.MouseButtonRelease):
        ev = QMouseEvent(etype, QPointF(50, 50), QPointF(50, 50), Qt.MouseButton.LeftButton,
                         Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        QApplication.sendEvent(card, ev)
    assert hits == [1]
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_widgets.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.ui.theme'`.

- [ ] **Paso 3: Implementar theme**

`knob/ui/theme.py`:
```python
# -*- coding: utf-8 -*-
"""Temas del estilo Glow, colores, tipografías y hoja de estilos global."""
from dataclasses import dataclass

from PySide6.QtGui import QColor, QFont, QFontDatabase

from knob.core.paths import resource_path

TEXT = "#F3EEFF"
GREEN, GREEN_TEXT = "#4ADE80", "#86EFAC"
AMBER, AMBER_TEXT = "#FBBF24", "#FCD34D"
RED, RED_TEXT = "#F87171", "#FCA5A5"
PANEL_BG = QColor(22, 16, 34, 247)

# Glifos de Segoe Fluent Icons (Win11) / Segoe MDL2 Assets (Win10)
GLYPH_VOLUME = ""
GLYPH_MIC = ""
GLYPH_FOCUS = ""
GLYPH_ADD = ""
GLYPH_SETTINGS = ""
GLYPH_LEFT = ""
GLYPH_RIGHT = ""
GLYPH_CLOSE = ""


@dataclass(frozen=True)
class Theme:
    name: str
    grad_a: str      # inicio del degradé (abajo / izquierda)
    grad_b: str      # fin del degradé
    halo: str        # halo radial del fondo
    bg: str          # fondo base y barra de título
    knob_face: str
    knob_ring: str
    pointer: str
    brand_a: str
    brand_b: str
    text_dim: str


THEMES = {
    "Nebula": Theme("Nebula", "#7C3AED", "#EC4899", "#2A1450", "#0C0A14", "#2B2238", "#4B3A6A", "#F0ABFC", "#C084FC", "#F472B6", "#B9AED6"),
    "Cobalt": Theme("Cobalt", "#2563EB", "#22D3EE", "#102A55", "#080B14", "#1D2638", "#33486A", "#7DD3FC", "#60A5FA", "#22D3EE", "#A9BEDB"),
    "Magma":  Theme("Magma",  "#DC2626", "#F97316", "#4A1410", "#120908", "#382222", "#6A3A33", "#FDBA74", "#F87171", "#FB923C", "#D9B3AE"),
    "Toxic":  Theme("Toxic",  "#059669", "#A3E635", "#0F3A26", "#08110C", "#1F3329", "#335A45", "#BEF264", "#34D399", "#A3E635", "#A8CDB8"),
    "Onyx":   Theme("Onyx",   "#737373", "#F5F5F5", "#2A2A2A", "#0C0C0C", "#2A2A2A", "#4A4A4A", "#FFFFFF", "#D4D4D4", "#FFFFFF", "#A3A3A3"),
    "Sunset": Theme("Sunset", "#DB2777", "#FBBF24", "#4A1636", "#120A0E", "#38222F", "#6A3A55", "#FDE68A", "#F472B6", "#FBBF24", "#D9B3C6"),
}
THEME_ORDER = ("Nebula", "Cobalt", "Magma", "Toxic", "Onyx", "Sunset")


def get_theme(name: str) -> Theme:
    return THEMES.get(name, THEMES["Nebula"])


def qcolor(hex_color: str, alpha: int = 255) -> QColor:
    c = QColor(hex_color)
    c.setAlpha(alpha)
    return c


def rgba_css(hex_color: str, alpha: int) -> str:
    c = QColor(hex_color)
    return f"rgba({c.red()},{c.green()},{c.blue()},{alpha})"


def ui_font(px: int, weight=QFont.Weight.Normal) -> QFont:
    f = QFont()
    f.setFamilies(["Segoe UI Variable Text", "Segoe UI"])
    f.setPixelSize(px)
    f.setWeight(weight)
    return f


def icon_font(px: int) -> QFont:
    f = QFont()
    f.setFamilies(["Segoe Fluent Icons", "Segoe MDL2 Assets"])
    f.setPixelSize(px)
    return f


_brand_family = None


def brand_font(px: int) -> QFont:
    """Fuente del logo (Amazing Views); si no se puede cargar, Segoe UI Black."""
    global _brand_family
    if _brand_family is None:
        fid = QFontDatabase.addApplicationFont(resource_path("fonts/AmazingViews.ttf"))
        families = QFontDatabase.applicationFontFamilies(fid) if fid >= 0 else []
        _brand_family = families[0] if families else ""
    f = QFont()
    if _brand_family:
        f.setFamilies([_brand_family])
    else:
        f.setFamilies(["Segoe UI Black", "Segoe UI"])
        f.setWeight(QFont.Weight.Black)
    f.setPixelSize(px)
    return f


def build_qss(t: Theme) -> str:
    """Hoja de estilos de los widgets estándar (los pintados a mano no la usan)."""
    sel = rgba_css(t.grad_a, 115)
    return f"""
    QWidget {{ color: {TEXT}; }}
    QLabel {{ background: transparent; }}
    QLabel#dim {{ color: {t.text_dim}; }}
    QLabel#section {{ color: {t.text_dim}; }}
    QLabel#error {{ color: {RED_TEXT}; }}
    QToolTip {{ background: #161022; color: {TEXT}; border: 1px solid rgba(255,255,255,26); padding: 4px 6px; }}
    QMenu {{ background: #161022; border: 1px solid rgba(255,255,255,26); border-radius: 8px; padding: 4px; }}
    QMenu::item {{ padding: 6px 22px 6px 12px; border-radius: 6px; }}
    QMenu::item:selected {{ background: {sel}; }}
    QMenu::indicator {{ width: 0px; }}
    QComboBox {{ background: rgba(255,255,255,15); border: 1px solid rgba(255,255,255,23); border-radius: 7px; padding: 5px 9px; }}
    QComboBox:hover {{ background: rgba(255,255,255,26); }}
    QComboBox::drop-down {{ border: none; width: 20px; }}
    QComboBox QAbstractItemView {{ background: #161022; border: 1px solid rgba(255,255,255,26); selection-background-color: {sel}; outline: none; padding: 4px; }}
    QLineEdit {{ background: rgba(255,255,255,15); border: 1px solid rgba(255,255,255,31); border-radius: 7px; padding: 7px 9px; selection-background-color: {sel}; }}
    QLineEdit:focus {{ border: 1px solid {t.brand_a}; }}
    QPushButton {{ background: rgba(255,255,255,18); border: 1px solid rgba(255,255,255,26); border-radius: 7px; padding: 6px 12px; }}
    QPushButton:hover {{ background: rgba(255,255,255,31); }}
    QPushButton:disabled {{ color: rgba(255,255,255,77); background: rgba(255,255,255,8); }}
    QPushButton#primary {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {rgba_css(t.grad_a, 200)}, stop:1 {rgba_css(t.grad_b, 150)}); border: none; font-weight: 600; }}
    QPushButton#primary:hover {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {t.grad_a}, stop:1 {rgba_css(t.grad_b, 200)}); }}
    QPushButton#danger {{ color: {RED_TEXT}; }}
    QPushButton#link {{ background: transparent; border: none; color: {t.text_dim}; padding: 4px 2px; text-align: left; }}
    QPushButton#link:hover {{ color: {TEXT}; }}
    QPushButton#profile {{ padding: 4px 10px; font-size: 12px; }}
    QToolButton {{ background: rgba(255,255,255,15); border: 1px solid rgba(255,255,255,20); border-radius: 7px; }}
    QToolButton:hover {{ background: rgba(255,255,255,31); }}
    QScrollArea {{ background: transparent; border: none; }}
    QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QScrollBar:vertical {{ background: transparent; width: 6px; margin: 0; }}
    QScrollBar::handle:vertical {{ background: rgba(255,255,255,38); border-radius: 3px; min-height: 24px; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    """
```

- [ ] **Paso 4: Implementar glass, fader y vu_meter**

`knob/ui/widgets/glass.py`:
```python
# -*- coding: utf-8 -*-
"""Superficies translúcidas del estilo Glow."""
from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QRadialGradient
from PySide6.QtWidgets import QWidget

from knob.ui.theme import qcolor


class GlassCard(QWidget):
    """Tarjeta translúcida con borde suave. Si es clickeable emite clicked."""
    clicked = Signal()

    def __init__(self, parent=None, clickable: bool = True):
        super().__init__(parent)
        self._clickable = clickable
        self._hover = False
        self._pressed = False
        self._dashed = False
        self._radius = 10
        if clickable:
            self.setCursor(Qt.CursorShape.PointingHandCursor)

    def set_dashed(self, dashed: bool) -> None:
        if dashed != self._dashed:
            self._dashed = dashed
            self.update()

    def enterEvent(self, e):
        self._hover = True
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hover = False
        self.update()
        super().leaveEvent(e)

    def mousePressEvent(self, e):
        if self._clickable and e.button() == Qt.MouseButton.LeftButton:
            self._pressed = True
            e.accept()
            return
        super().mousePressEvent(e)

    def mouseReleaseEvent(self, e):
        if self._pressed and e.button() == Qt.MouseButton.LeftButton:
            self._pressed = False
            if self.rect().contains(e.position().toPoint()):
                self.clicked.emit()
            e.accept()
            return
        super().mouseReleaseEvent(e)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        hover = self._hover and self._clickable
        if self._dashed:
            p.setPen(QPen(QColor(255, 255, 255, 60 if hover else 46), 1, Qt.PenStyle.DashLine))
            p.setBrush(QColor(255, 255, 255, 10) if hover else Qt.BrushStyle.NoBrush)
        else:
            p.setPen(QPen(QColor(255, 255, 255, 30 if hover else 18), 1))
            p.setBrush(QColor(255, 255, 255, 17 if hover else 10))
        p.drawRoundedRect(r, self._radius, self._radius)


class Backdrop(QWidget):
    """Fondo de la ventana: color base + halo radial arriba a la derecha."""

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(self._t.bg))
        w = self.width()
        g = QRadialGradient(QPointF(w * 0.78, -20), w * 0.95)
        g.setColorAt(0.0, qcolor(self._t.halo, 255))
        g.setColorAt(0.55, qcolor(self._t.halo, 70))
        g.setColorAt(1.0, qcolor(self._t.bg, 0))
        p.fillRect(self.rect(), g)
```

`knob/ui/widgets/fader.py`:
```python
# -*- coding: utf-8 -*-
"""Fader vertical de solo lectura: pista, relleno con degradé y brillo, tapa blanca."""
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QLinearGradient, QPainter
from PySide6.QtWidgets import QWidget

from knob.ui.theme import qcolor

TRACK_W = 8
CAP_W = 22
CAP_H = 10
V_MARGIN = CAP_H / 2 + 2      # VuMeter usa el mismo margen para alinearse


class Fader(QWidget):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self._v = 0.0
        self._muted = False
        self._active = True
        self.setFixedWidth(28)
        self.setMinimumHeight(100)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_value(self, v: float) -> None:
        v = max(0.0, min(1.0, v))
        if abs(v - self._v) > 0.0005:
            self._v = v
            self.update()

    def set_muted(self, muted: bool) -> None:
        if muted != self._muted:
            self._muted = muted
            self.update()

    def set_active(self, active: bool) -> None:
        if active != self._active:
            self._active = active
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        cx = self.width() / 2
        top, bottom = V_MARGIN, self.height() - V_MARGIN
        usable = bottom - top
        track = QRectF(cx - TRACK_W / 2, top, TRACK_W, usable)
        p.setBrush(QColor(255, 255, 255, 20))
        p.drawRoundedRect(track, TRACK_W / 2, TRACK_W / 2)

        y = bottom - self._v * usable
        lit = self._active and not self._muted
        if self._active:
            fill = QRectF(track.left(), y, TRACK_W, bottom - y)
            if fill.height() > 0.5:
                if lit:
                    for grow, alpha in ((5.0, 26), (2.5, 46)):          # brillo
                        p.setBrush(qcolor(self._t.grad_a, alpha))
                        p.drawRoundedRect(fill.adjusted(-grow, -grow, grow, grow),
                                          TRACK_W / 2 + grow, TRACK_W / 2 + grow)
                    g = QLinearGradient(0, bottom, 0, top)
                    g.setColorAt(0.0, QColor(self._t.grad_a))
                    g.setColorAt(1.0, QColor(self._t.grad_b))
                    p.setBrush(g)
                else:
                    p.setBrush(QColor(255, 255, 255, 45))
                p.drawRoundedRect(fill, TRACK_W / 2, TRACK_W / 2)

        cap = QRectF(cx - CAP_W / 2, y - CAP_H / 2, CAP_W, CAP_H)
        if lit:
            p.setBrush(qcolor(self._t.brand_a, 70))
            p.drawRoundedRect(cap.adjusted(-3, -3, 3, 3), 8, 8)
            p.setBrush(QColor("#FFFFFF"))
        else:
            p.setBrush(QColor(255, 255, 255, 90))
        p.drawRoundedRect(cap, CAP_H / 2, CAP_H / 2)
```

`knob/ui/widgets/vu_meter.py`:
```python
# -*- coding: utf-8 -*-
"""Medidor de nivel fino con retención de pico (verde / ámbar / rojo como v3.3)."""
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QWidget

from knob.ui.theme import AMBER, GREEN, RED
from knob.ui.widgets.fader import V_MARGIN

BAR_W = 3


class VuMeter(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._level = 0.0
        self._peak = 0.0
        self._active = True
        self.setFixedWidth(5)
        self.setMinimumHeight(100)

    def set_level(self, level: float, peak: float) -> None:
        if abs(level - self._level) > 0.002 or abs(peak - self._peak) > 0.002:
            self._level, self._peak = level, peak
            self.update()

    def set_active(self, active: bool) -> None:
        if active != self._active:
            self._active = active
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        x = (self.width() - BAR_W) / 2
        top, bottom = V_MARGIN, self.height() - V_MARGIN
        usable = bottom - top
        p.setBrush(QColor(255, 255, 255, 15))
        p.drawRoundedRect(QRectF(x, top, BAR_W, usable), 1.5, 1.5)
        if not self._active:
            return
        h = self._level * usable
        if h > 0.5:
            color = RED if self._level > 0.85 else AMBER if self._level > 0.70 else GREEN
            p.setBrush(QColor(color))
            p.drawRoundedRect(QRectF(x, bottom - h, BAR_W, h), 1.5, 1.5)
        if self._peak > 0.01:
            p.setBrush(QColor(255, 255, 255, 170))
            p.drawRect(QRectF(x, bottom - self._peak * usable - 1, BAR_W, 2))
```

- [ ] **Paso 5: Implementar knob_dial, brand, status_pill y toggle**

`knob/ui/widgets/knob_dial.py`:
```python
# -*- coding: utf-8 -*-
"""Perilla estilo hardware: 5 marcas de escala, cara oscura y puntero luminoso (barrido de 270°)."""
import math

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QWidget

from knob.ui.theme import qcolor

SIZE = 44
_TICKS = (-135.0, -67.5, 0.0, 67.5, 135.0)   # grados desde las 12, sentido horario


def _polar(cx, cy, r, deg):
    a = math.radians(deg)
    return QPointF(cx + r * math.sin(a), cy - r * math.cos(a))


class KnobDial(QWidget):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self._v = 0.0
        self._muted = False
        self._active = True
        self.setFixedSize(SIZE, SIZE)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_value(self, v: float) -> None:
        v = max(0.0, min(1.0, v))
        if abs(v - self._v) > 0.001:
            self._v = v
            self.update()

    def set_muted(self, muted: bool) -> None:
        if muted != self._muted:
            self._muted = muted
            self.update()

    def set_active(self, active: bool) -> None:
        if active != self._active:
            self._active = active
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        cx, cy = self.width() / 2, self.height() / 2
        p.setPen(QPen(QColor(self._t.knob_ring), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        for deg in _TICKS:
            p.drawLine(_polar(cx, cy, 17.5, deg), _polar(cx, cy, 20.5, deg))
        p.setPen(QPen(QColor(self._t.knob_ring), 1))
        p.setBrush(QColor(self._t.knob_face))
        p.drawEllipse(QPointF(cx, cy), 13, 13)

        angle = -135.0 + self._v * 270.0
        start, end = _polar(cx, cy, 2.5, angle), _polar(cx, cy, 10.5, angle)
        lit = self._active and not self._muted
        if lit:
            p.setPen(QPen(qcolor(self._t.pointer, 70), 5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
            p.drawLine(start, end)
            color = QColor(self._t.pointer)
        else:
            color = qcolor(self._t.text_dim, 120)
        p.setPen(QPen(color, 2.2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        p.drawLine(start, end)
```

`knob/ui/widgets/brand.py`:
```python
# -*- coding: utf-8 -*-
"""Logo 'KNOB' con la fuente Amazing Views rellena con el degradé del tema."""
import math

from PySide6.QtGui import QColor, QFontMetricsF, QLinearGradient, QPainter, QPainterPath
from PySide6.QtWidgets import QWidget

from knob.ui.theme import brand_font

TEXT = "KNOB"


class BrandLabel(QWidget):
    def __init__(self, theme, px: int = 26, parent=None):
        super().__init__(parent)
        self._t = theme
        self._font = brand_font(px)
        fm = QFontMetricsF(self._font)
        self._ascent = fm.ascent()
        self.setFixedSize(math.ceil(fm.horizontalAdvance(TEXT)) + 4, math.ceil(fm.height()))

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.addText(2, self._ascent, self._font, TEXT)
        g = QLinearGradient(0, 0, self.width(), 0)
        g.setColorAt(0.0, QColor(self._t.brand_a))
        g.setColorAt(1.0, QColor(self._t.brand_b))
        p.fillPath(path, g)
```

`knob/ui/widgets/status_pill.py`:
```python
# -*- coding: utf-8 -*-
"""Pastilla de conexión: verde conectado, ámbar buscando (punto que late), rojo puerto ocupado."""
import math

from PySide6.QtCore import QPointF, QRectF, Qt, QTimer
from PySide6.QtGui import QColor, QFontMetrics, QPainter
from PySide6.QtWidgets import QWidget

from knob.ui.theme import AMBER, AMBER_TEXT, GREEN, GREEN_TEXT, RED, RED_TEXT, qcolor, ui_font

_STYLES = {"connected": (GREEN, GREEN_TEXT), "searching": (AMBER, AMBER_TEXT), "busy": (RED, RED_TEXT)}


class StatusPill(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._font = ui_font(11)
        self._state = ""
        self._text = ""
        self._phase = 0.0
        self._timer = QTimer(self)
        self._timer.setInterval(50)
        self._timer.timeout.connect(self._pulse)
        self.set_state("searching", "Buscando KNOB…")

    def state(self) -> str:
        return self._state

    def set_state(self, state: str, text: str) -> None:
        if state == self._state and text == self._text:
            return
        self._state, self._text = state, text
        fm = QFontMetrics(self._font)
        self.setFixedSize(fm.horizontalAdvance(text) + 30, 22)
        self.setToolTip(text)
        if state == "searching":
            self._timer.start()
        else:
            self._timer.stop()
        self.update()

    def _pulse(self) -> None:
        self._phase = (self._phase + 0.04) % 1.0
        self.update()

    def paintEvent(self, _):
        dot, text_color = _STYLES.get(self._state, _STYLES["searching"])
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(qcolor(dot, 32))
        p.drawRoundedRect(QRectF(self.rect()), 11, 11)
        alpha = 255
        if self._state == "searching":
            alpha = int(110 + 145 * (0.5 + 0.5 * math.cos(2 * math.pi * self._phase)))
        center = QPointF(12, self.height() / 2)
        p.setBrush(qcolor(dot, alpha // 4))
        p.drawEllipse(center, 5.5, 5.5)
        p.setBrush(qcolor(dot, alpha))
        p.drawEllipse(center, 3, 3)
        p.setPen(QColor(text_color))
        p.setFont(self._font)
        p.drawText(QRectF(21, 0, self.width() - 21, self.height()),
                   Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, self._text)
```

`knob/ui/widgets/toggle.py`:
```python
# -*- coding: utf-8 -*-
"""Interruptor on/off con el degradé del tema."""
from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QLinearGradient, QPainter
from PySide6.QtWidgets import QAbstractButton


class Toggle(QAbstractButton):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self.setCheckable(True)
        self.setFixedSize(34, 19)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def sizeHint(self):
        return QSize(34, 19)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        r = QRectF(self.rect())
        if self.isChecked():
            g = QLinearGradient(0, 0, r.width(), 0)
            g.setColorAt(0, QColor(self._t.grad_a))
            g.setColorAt(1, QColor(self._t.grad_b))
            p.setBrush(g)
        else:
            p.setBrush(QColor(255, 255, 255, 38))
        p.drawRoundedRect(r, r.height() / 2, r.height() / 2)
        knob_x = r.width() - r.height() / 2 if self.isChecked() else r.height() / 2
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(QPointF(knob_x, r.height() / 2), 7, 7)
```

- [ ] **Paso 6: Implementar icon_badge**

`knob/ui/widgets/icon_badge.py`:
```python
# -*- coding: utf-8 -*-
"""Ícono de un destino: ícono real del exe, inicial de color, glifo de sistema o '+' (vacío)."""
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QWidget

from knob.ui.theme import (GLYPH_ADD, GLYPH_FOCUS, GLYPH_MIC, GLYPH_VOLUME, get_theme, icon_font,
                           qcolor, ui_font)

_SYSTEM_GLYPHS = {"master": GLYPH_VOLUME, "mic": GLYPH_MIC, "focus": GLYPH_FOCUS}
_LETTER_COLORS = ("#D6293E", "#5865F2", "#1DB954", "#E5A50A", "#0EA5E9", "#A855F7", "#F97316", "#14B8A6")


def letter_color(name: str) -> str:
    return _LETTER_COLORS[sum(map(ord, name)) % len(_LETTER_COLORS)] if name else _LETTER_COLORS[0]


class IconBadge(QWidget):
    def __init__(self, size: int = 28, parent=None):
        super().__init__(parent)
        self._size = size
        self._name = ""
        self._pixmap = None
        self._kind = "unassigned"
        self._dimmed = False
        self._t = get_theme("Nebula")
        self.setFixedSize(size, size)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_content(self, name: str, pixmap, kind) -> None:
        self._name, self._pixmap, self._kind = name or "", pixmap, kind
        self.update()

    def set_dimmed(self, dimmed: bool) -> None:
        if dimmed != self._dimmed:
            self._dimmed = dimmed
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        if self._dimmed:
            p.setOpacity(0.4)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        radius = self._size * 0.28
        glyph_px = max(10, int(self._size * 0.52))
        if self._kind == "unassigned":
            p.setPen(QPen(QColor(255, 255, 255, 77), 1, Qt.PenStyle.DashLine))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(r, radius, radius)
            p.setPen(QColor(self._t.text_dim))
            p.setFont(icon_font(glyph_px - 2))
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, GLYPH_ADD)
        elif self._kind in _SYSTEM_GLYPHS:
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(qcolor(self._t.grad_a, 90))
            p.drawRoundedRect(r, radius, radius)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(icon_font(glyph_px))
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, _SYSTEM_GLYPHS[self._kind])
        elif self._pixmap is not None and not self._pixmap.isNull():
            p.drawPixmap(r.toRect(), self._pixmap)
        else:
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(letter_color(self._name)))
            p.drawRoundedRect(r, radius, radius)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(ui_font(int(self._size * 0.5), QFont.Weight.Bold))
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, (self._name[:1] or "?").upper())
```

- [ ] **Paso 7: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest tests/test_widgets.py`
Esperado: `6 passed`.

- [ ] **Paso 8: Commit**

```bash
git add knob/ui/theme.py knob/ui/widgets tests/test_widgets.py
git commit -m "feat: temas Glow y widgets base (fader, VU, perilla, logo, estado, interruptor, íconos)" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 9: Tarjetas, ventana principal y vista previa

**Archivos:**
- Crear: `knob/ui/widgets/channel_card.py`, `knob/ui/widgets/encoder_card.py`, `knob/ui/widgets/output_bar.py`, `knob/ui/icons.py`, `knob/ui/main_window.py`, `knob/ui/preview.py`
- Test: `tests/test_main_window.py`

**Interfaces:**
- Consume: widgets y theme de la Tarea 8; `system.style_title_bar`.
- Produce:
  - `ChannelCard(index, theme)` y `EncoderCard(index, theme)`, ambos `GlassCard` (señal `clicked`), con:
    - `set_assignment(title: str|None, pixmap: QPixmap|None, kind: str|None)`.
    - `set_value(raw: float)` (0..1023), `set_muted(bool)`, `set_focus_target(name: str|None)`, `set_theme(t)`.
    - Solo `ChannelCard`: `set_level(level, peak)`. Solo `EncoderCard`: señal `gear_clicked`.
  - `OutputBar(theme)` con señales `prev_clicked` y `next_clicked`, `set_name(name: str|None, flash: bool=False)` y `set_theme(t)`.
  - `AppIcons(exe_lookup: Callable[[str], str|None])` con `icon_for(app_name) -> QPixmap | None`.
  - `MainWindow(theme)` con:
    - Señales `channel_clicked(int)` (0..4), `encoder_gear_clicked(int)` (0..1), `settings_clicked()`, `profile_selected(str)`, `output_prev()`, `output_next()` y `close_requested()`.
    - Atributos `status: StatusPill`, `output_bar: OutputBar`, `channel_cards` (3) y `encoder_cards` (2).
    - Métodos `card_for(ch)`, `set_profiles(names, active)`, `set_update_badge(bool)` y `apply_theme(t)`.
  - `python -m knob.ui.preview [Tema] [--screenshot ruta.png]`.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_main_window.py`:
```python
# -*- coding: utf-8 -*-
import pytest


@pytest.mark.parametrize("theme_name", ["Nebula", "Cobalt", "Magma", "Toxic", "Onyx", "Sunset"])
def test_ventana_se_construye_y_pinta(qapp, theme_name):
    from knob.ui.main_window import MainWindow
    from knob.ui.theme import build_qss, get_theme
    t = get_theme(theme_name)
    qapp.setStyleSheet(build_qss(t))
    w = MainWindow(t)
    w.apply_theme(t)
    w.set_profiles(["Default", "Gaming"], "Gaming")
    w.card_for(0).set_assignment("Opera", None, None)
    w.card_for(1).set_assignment(None, None, None)
    w.card_for(2).set_assignment("App en foco", None, "focus")
    w.card_for(2).set_focus_target("Chrome")
    w.card_for(3).set_assignment("Volumen general", None, "master")
    w.card_for(4).set_assignment(None, None, None)
    for ch in range(5):
        w.card_for(ch).set_value(700)
    w.card_for(0).set_muted(True)
    w.channel_cards[0].set_level(0.6, 0.8)
    w.output_bar.set_name("Realtek Digital Output", flash=True)
    w.status.set_state("connected", "Conectado")
    w.set_update_badge(True)
    assert w.width() == 520 and w.height() == 620
    assert not w.grab().isNull()


def test_senales_de_tarjetas(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.theme import get_theme
    w = MainWindow(get_theme("Nebula"))
    got = []
    w.channel_clicked.connect(lambda ch: got.append(("ch", ch)))
    w.encoder_gear_clicked.connect(lambda e: got.append(("gear", e)))
    w.profile_selected.connect(lambda n: got.append(("profile", n)))
    w.channel_cards[1].clicked.emit()
    w.encoder_cards[0].clicked.emit()
    w.encoder_cards[1].gear_clicked.emit()
    w.set_profiles(["Default", "Gaming"], "Default")
    w._profile_menu.actions()[1].trigger()
    assert got == [("ch", 1), ("ch", 3), ("gear", 1), ("profile", "Gaming")]


def test_close_emite_close_requested(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.theme import get_theme
    w = MainWindow(get_theme("Nebula"))
    hits = []
    w.close_requested.connect(lambda: hits.append(1))
    w.show()
    w.close()
    assert hits == [1]
    assert w.isVisible()          # la ventana no se cierra sola: decide el controlador


def test_canal_muestra_porcentaje_y_mute(qapp):
    from knob.ui.theme import get_theme
    from knob.ui.widgets.channel_card import ChannelCard
    card = ChannelCard(0, get_theme("Nebula"))
    card.set_assignment("Discord", None, None)
    card.set_value(1023)
    assert card.value_label.text() == "100%"
    card.set_muted(True)
    assert card.value_label.text() == "MUTE"
    card.set_muted(False)
    card.set_value(512)
    assert card.value_label.text() == "50%"
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_main_window.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.ui.main_window'`.

- [ ] **Paso 3: Implementar channel_card y encoder_card**

`knob/ui/widgets/channel_card.py`:
```python
# -*- coding: utf-8 -*-
"""Tarjeta de canal (faders 1-3): ícono, nombre, fader + VU y porcentaje."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontMetrics
from PySide6.QtWidgets import QHBoxLayout, QLabel, QVBoxLayout

from knob.ui.theme import RED_TEXT, ui_font
from knob.ui.widgets.fader import Fader
from knob.ui.widgets.glass import GlassCard
from knob.ui.widgets.icon_badge import IconBadge
from knob.ui.widgets.vu_meter import VuMeter

NAME_MAX_PX = 120


class ChannelCard(GlassCard):
    def __init__(self, index: int, theme, parent=None):
        super().__init__(parent)
        self.index = index
        self._t = theme
        self._assigned = False
        self._muted = False
        self._kind = None
        self._pct = None
        self.setToolTip("Click para elegir qué controla este fader")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 14, 8, 12)
        lay.setSpacing(6)
        self.badge = IconBadge(28)
        lay.addWidget(self.badge, 0, Qt.AlignmentFlag.AlignHCenter)

        self.name_label = QLabel("Sin asignar")
        self.name_label.setObjectName("dim")
        self.name_label.setFont(ui_font(12))
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.name_label)

        self.sub_label = QLabel("")
        self.sub_label.setObjectName("dim")
        self.sub_label.setFont(ui_font(10))
        self.sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sub_label.hide()
        lay.addWidget(self.sub_label)

        row = QHBoxLayout()
        row.setSpacing(3)
        row.addStretch(1)
        self.fader = Fader(theme)
        self.vu = VuMeter()
        row.addWidget(self.fader)
        row.addWidget(self.vu)
        row.addStretch(1)
        lay.addLayout(row, 1)

        self.value_label = QLabel("")
        self.value_label.setFont(ui_font(14, QFont.Weight.DemiBold))
        self.value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.value_label)
        self.set_assignment(None, None, None)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.fader.set_theme(theme)
        self.badge.set_theme(theme)

    def set_assignment(self, title, pixmap, kind) -> None:
        self._assigned = title is not None
        self._kind = kind if self._assigned else None
        self.set_dashed(not self._assigned)
        self.fader.set_active(self._assigned)
        self.vu.set_active(self._assigned)
        if not self._assigned:
            self.badge.set_content("", None, "unassigned")
            self.name_label.setText("Sin asignar")
            self.sub_label.setText("click para elegir")
            self.sub_label.show()
            self.value_label.setText("")
            self._pct = None
            return
        self.badge.set_content(title, pixmap, kind)
        fm = QFontMetrics(self.name_label.font())
        self.name_label.setText(fm.elidedText(title, Qt.TextElideMode.ElideRight, NAME_MAX_PX))
        self.name_label.setToolTip(title)
        if kind != "focus":
            self.sub_label.hide()
        self._pct = None
        self._render_value()

    def set_focus_target(self, name) -> None:
        if self._kind != "focus":
            return
        self.sub_label.setText(f"→ {name}" if name else "→ ninguna")
        self.sub_label.show()

    def set_value(self, raw: float) -> None:
        self.fader.set_value(raw / 1023.0)
        pct = round(raw / 1023.0 * 100)
        if pct != self._pct:
            self._pct = pct
            self._render_value()

    def set_muted(self, muted: bool) -> None:
        if muted == self._muted:
            return
        self._muted = muted
        self.fader.set_muted(muted)
        self.badge.set_dimmed(muted)
        self._render_value()

    def set_level(self, level: float, peak: float) -> None:
        self.vu.set_level(level if self._assigned else 0.0, peak if self._assigned else 0.0)

    def _render_value(self) -> None:
        if not self._assigned:
            return
        if self._muted:
            self.value_label.setText("MUTE")
            self.value_label.setStyleSheet(f"color: {RED_TEXT};")
        else:
            self.value_label.setText(f"{self._pct or 0}%")
            self.value_label.setStyleSheet("")
```

`knob/ui/widgets/encoder_card.py`:
```python
# -*- coding: utf-8 -*-
"""Tarjeta de encoder: perilla, destino del giro y engranaje para las acciones del botón."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QFontMetrics
from PySide6.QtWidgets import QHBoxLayout, QLabel, QToolButton, QVBoxLayout

from knob.ui.theme import GLYPH_SETTINGS, RED_TEXT, icon_font, ui_font
from knob.ui.widgets.glass import GlassCard
from knob.ui.widgets.icon_badge import IconBadge
from knob.ui.widgets.knob_dial import KnobDial

CAPTIONS = ("Encoder izq.", "Encoder der.")
TITLE_MAX_PX = 120


class EncoderCard(GlassCard):
    gear_clicked = Signal()

    def __init__(self, index: int, theme, parent=None):
        super().__init__(parent)
        self.index = index
        self._t = theme
        self._assigned = False
        self._muted = False
        self._kind = None
        self.setToolTip("Click para elegir qué controla este encoder")

        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 9, 9, 9)
        lay.setSpacing(10)
        self.dial = KnobDial(theme)
        lay.addWidget(self.dial)

        col = QVBoxLayout()
        col.setSpacing(1)
        self.caption = QLabel(CAPTIONS[index])
        self.caption.setObjectName("dim")
        self.caption.setFont(ui_font(10))
        col.addWidget(self.caption)
        title_row = QHBoxLayout()
        title_row.setSpacing(5)
        self.badge = IconBadge(16)
        title_row.addWidget(self.badge)
        self.title_label = QLabel("Sin asignar")
        self.title_label.setFont(ui_font(13, QFont.Weight.DemiBold))
        title_row.addWidget(self.title_label, 1)
        col.addLayout(title_row)
        self.sub_label = QLabel("")
        self.sub_label.setObjectName("dim")
        self.sub_label.setFont(ui_font(10))
        self.sub_label.hide()
        col.addWidget(self.sub_label)
        lay.addLayout(col, 1)

        self.gear = QToolButton()
        self.gear.setFont(icon_font(13))
        self.gear.setText(GLYPH_SETTINGS)
        self.gear.setFixedSize(26, 26)
        self.gear.setCursor(Qt.CursorShape.PointingHandCursor)
        self.gear.setToolTip("Acciones del botón (click, doble click, mantener)")
        self.gear.clicked.connect(self.gear_clicked)
        lay.addWidget(self.gear, 0, Qt.AlignmentFlag.AlignTop)
        self.set_assignment(None, None, None)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.dial.set_theme(theme)
        self.badge.set_theme(theme)

    def set_assignment(self, title, pixmap, kind) -> None:
        self._assigned = title is not None
        self._kind = kind if self._assigned else None
        self.set_dashed(not self._assigned)
        self.dial.set_active(self._assigned)
        if not self._assigned:
            self.badge.set_content("", None, "unassigned")
            self.title_label.setText("Sin asignar")
            self.sub_label.hide()
            return
        self.badge.set_content(title, pixmap, kind)
        fm = QFontMetrics(self.title_label.font())
        self.title_label.setText(fm.elidedText(title, Qt.TextElideMode.ElideRight, TITLE_MAX_PX))
        self.title_label.setToolTip(title)
        if kind != "focus":
            self.sub_label.hide()

    def set_focus_target(self, name) -> None:
        if self._kind != "focus":
            return
        self.sub_label.setText(f"→ {name}" if name else "→ ninguna")
        self.sub_label.show()

    def set_value(self, raw: float) -> None:
        self.dial.set_value(raw / 1023.0)

    def set_muted(self, muted: bool) -> None:
        if muted == self._muted:
            return
        self._muted = muted
        self.dial.set_muted(muted)
        self.badge.set_dimmed(muted)
        self.title_label.setStyleSheet(f"color: {RED_TEXT};" if muted else "")
```

- [ ] **Paso 4: Implementar output_bar e icons**

`knob/ui/widgets/output_bar.py`:
```python
# -*- coding: utf-8 -*-
"""Barra inferior: salida de audio predeterminada con flechas para cambiarla."""
from PySide6.QtCore import QRectF, Qt, QVariantAnimation, Signal
from PySide6.QtGui import QPainter, QPen
from PySide6.QtWidgets import QHBoxLayout, QLabel, QToolButton

from knob.ui.theme import GLYPH_LEFT, GLYPH_RIGHT, GLYPH_VOLUME, icon_font, qcolor, ui_font
from knob.ui.widgets.glass import GlassCard


class OutputBar(GlassCard):
    prev_clicked = Signal()
    next_clicked = Signal()

    def __init__(self, theme, parent=None):
        super().__init__(parent, clickable=False)
        self._t = theme
        self._flash = 0.0
        self.setFixedHeight(40)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 5, 6, 5)
        lay.setSpacing(8)
        icon = QLabel(GLYPH_VOLUME)
        icon.setFont(icon_font(14))
        icon.setObjectName("dim")
        lay.addWidget(icon)
        self.name_label = QLabel("")
        self.name_label.setFont(ui_font(12))
        lay.addWidget(self.name_label, 1)
        for glyph, signal, tip in ((GLYPH_LEFT, self.prev_clicked, "Salida anterior"),
                                   (GLYPH_RIGHT, self.next_clicked, "Salida siguiente")):
            b = QToolButton()
            b.setFont(icon_font(11))
            b.setText(glyph)
            b.setFixedSize(28, 28)
            b.setToolTip(tip)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.clicked.connect(signal)
            lay.addWidget(b)
        self._anim = QVariantAnimation(self)
        self._anim.setStartValue(1.0)
        self._anim.setEndValue(0.0)
        self._anim.setDuration(1500)
        self._anim.valueChanged.connect(self._on_flash)
        self.set_name(None)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_name(self, name, flash: bool = False) -> None:
        self.name_label.setText(name or "Sin dispositivo de salida")
        self.name_label.setToolTip(name or "")
        if flash:
            self._anim.stop()
            self._anim.start()

    def _on_flash(self, value) -> None:
        self._flash = float(value)
        self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if self._flash <= 0.01:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(QPen(qcolor(self._t.brand_a, int(200 * self._flash)), 1.5))
        p.setBrush(qcolor(self._t.grad_a, int(50 * self._flash)))
        p.drawRoundedRect(QRectF(self.rect()).adjusted(1, 1, -1, -1), 10, 10)
```

`knob/ui/icons.py`:
```python
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
```

- [ ] **Paso 5: Implementar main_window**

`knob/ui/main_window.py`:
```python
# -*- coding: utf-8 -*-
"""Ventana única de KNOB: barra superior, 3 canales, 2 encoders y salida de audio."""
from PySide6.QtCore import QPoint, QPointF, Qt, Signal
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QHBoxLayout, QMainWindow, QMenu, QPushButton, QToolButton, QVBoxLayout

from knob.core import system
from knob.ui.theme import GLYPH_SETTINGS, icon_font
from knob.ui.widgets.brand import BrandLabel
from knob.ui.widgets.channel_card import ChannelCard
from knob.ui.widgets.encoder_card import EncoderCard
from knob.ui.widgets.glass import Backdrop
from knob.ui.widgets.output_bar import OutputBar
from knob.ui.widgets.status_pill import StatusPill

WINDOW_W, WINDOW_H = 520, 620


class GearButton(QToolButton):
    """Engranaje de ajustes con un punto cuando hay actualización."""

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self._badge = False
        self.setFont(icon_font(14))
        self.setText(GLYPH_SETTINGS)
        self.setFixedSize(30, 30)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("Ajustes")

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_badge(self, on: bool) -> None:
        self._badge = on
        self.setToolTip("Ajustes · hay una actualización disponible" if on else "Ajustes")
        self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if self._badge:
            p = QPainter(self)
            p.setRenderHint(QPainter.RenderHint.Antialiasing)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(self._t.grad_b))
            p.drawEllipse(QPointF(self.width() - 6, 6), 3.5, 3.5)


class MainWindow(QMainWindow):
    channel_clicked = Signal(int)
    encoder_gear_clicked = Signal(int)
    settings_clicked = Signal()
    profile_selected = Signal(str)
    output_prev = Signal()
    output_next = Signal()
    close_requested = Signal()

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self.setWindowTitle("KNOB")
        self.setFixedSize(WINDOW_W, WINDOW_H)

        self.backdrop = Backdrop(theme)
        self.setCentralWidget(self.backdrop)
        root = QVBoxLayout(self.backdrop)
        root.setContentsMargins(16, 14, 16, 16)
        root.setSpacing(12)

        top = QHBoxLayout()
        top.setSpacing(10)
        self.brand = BrandLabel(theme)
        top.addWidget(self.brand, 0, Qt.AlignmentFlag.AlignVCenter)
        self.status = StatusPill()
        top.addWidget(self.status, 0, Qt.AlignmentFlag.AlignVCenter)
        top.addStretch(1)
        self.profile_btn = QPushButton("Perfil")
        self.profile_btn.setObjectName("profile")
        self.profile_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._profile_menu = QMenu(self)
        self.profile_btn.clicked.connect(self._show_profile_menu)
        top.addWidget(self.profile_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        self.gear_btn = GearButton(theme)
        self.gear_btn.clicked.connect(self.settings_clicked)
        top.addWidget(self.gear_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        root.addLayout(top)

        chans = QHBoxLayout()
        chans.setSpacing(10)
        self.channel_cards = []
        for i in range(3):
            card = ChannelCard(i, theme)
            card.clicked.connect(lambda i=i: self.channel_clicked.emit(i))
            chans.addWidget(card)
            self.channel_cards.append(card)
        root.addLayout(chans, 1)

        encs = QHBoxLayout()
        encs.setSpacing(10)
        self.encoder_cards = []
        for i in range(2):
            card = EncoderCard(i, theme)
            card.clicked.connect(lambda i=i: self.channel_clicked.emit(3 + i))
            card.gear_clicked.connect(lambda i=i: self.encoder_gear_clicked.emit(i))
            encs.addWidget(card)
            self.encoder_cards.append(card)
        root.addLayout(encs)

        self.output_bar = OutputBar(theme)
        self.output_bar.prev_clicked.connect(self.output_prev)
        self.output_bar.next_clicked.connect(self.output_next)
        root.addWidget(self.output_bar)

    # --- API para el controlador ---------------------------------------
    def card_for(self, ch: int):
        return (self.channel_cards + self.encoder_cards)[ch]

    def set_profiles(self, names, active: str) -> None:
        self.profile_btn.setText(f"Perfil: {active}  ▾")
        self._profile_menu.clear()
        for name in names:
            act = self._profile_menu.addAction(("✓  " if name == active else "     ") + name)
            act.triggered.connect(lambda _=False, n=name: self.profile_selected.emit(n))

    def set_update_badge(self, on: bool) -> None:
        self.gear_btn.set_badge(on)

    def apply_theme(self, theme) -> None:
        self._t = theme
        self.backdrop.set_theme(theme)
        self.brand.set_theme(theme)
        self.gear_btn.set_theme(theme)
        self.output_bar.set_theme(theme)
        for card in self.channel_cards + self.encoder_cards:
            card.set_theme(theme)
        self._style_title_bar()

    # --- internos ------------------------------------------------------
    def _show_profile_menu(self) -> None:
        self._profile_menu.exec(self.profile_btn.mapToGlobal(QPoint(0, self.profile_btn.height() + 4)))

    def _style_title_bar(self) -> None:
        if self.isVisible():
            system.style_title_bar(int(self.winId()), self._t.bg)

    def showEvent(self, e):
        super().showEvent(e)
        self._style_title_bar()

    def closeEvent(self, e):
        e.ignore()                     # el controlador decide: bandeja o salir
        self.close_requested.emit()
```

- [ ] **Paso 6: Implementar la vista previa**

`knob/ui/preview.py`:
```python
# -*- coding: utf-8 -*-
"""Vista previa de la interfaz con datos falsos animados (sin hardware ni audio).

Uso:
  python -m knob.ui.preview [Tema] [--screenshot ruta.png]
"""
import math
import sys

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from knob.ui.main_window import MainWindow
from knob.ui.theme import build_qss, get_theme, ui_font


def build_demo_window(theme):
    w = MainWindow(theme)
    w.apply_theme(theme)
    w.set_profiles(["Default", "Gaming", "Trabajo"], "Gaming")
    w.status.set_state("connected", "Conectado")
    w.output_bar.set_name("Realtek Digital Output (Realtek(R) Audio)")
    w.card_for(0).set_assignment("Opera", None, None)
    w.card_for(1).set_assignment("Discord", None, None)
    w.card_for(2).set_assignment("Spotify", None, None)
    w.card_for(3).set_assignment("Volumen general", None, "master")
    w.card_for(4).set_assignment("App en foco", None, "focus")
    w.card_for(4).set_focus_target("Chrome")
    for ch, raw in enumerate((490, 860, 870, 715, 330)):
        w.card_for(ch).set_value(raw)
    for i, lvl in enumerate((0.45, 0.72, 0.6)):
        w.channel_cards[i].set_level(lvl, lvl + 0.1)
    return w


def main(argv):
    args = [a for a in argv[1:] if not a.startswith("--")]
    shot = argv[argv.index("--screenshot") + 1] if "--screenshot" in argv else None
    if shot in args:
        args.remove(shot)
    app = QApplication(argv)
    app.setStyle("Fusion")
    app.setFont(ui_font(13))
    theme = get_theme(args[0] if args else "Nebula")
    app.setStyleSheet(build_qss(theme))
    w = build_demo_window(theme)
    w.close_requested.connect(app.quit)
    w.show()

    t = {"v": 0.0}

    def animate():
        t["v"] += 0.05
        for i in range(3):
            raw = 512 + 400 * math.sin(t["v"] + i)
            w.card_for(i).set_value(raw)
            lvl = 0.5 + 0.4 * math.sin(t["v"] * 3 + i * 2)
            w.channel_cards[i].set_level(lvl, min(1.0, lvl + 0.08))
        w.card_for(3).set_value(512 + 400 * math.sin(t["v"] * 0.7))

    if shot:
        def capture():
            w.grab().save(shot)
            print(f"Captura guardada en {shot}")
            app.quit()
        QTimer.singleShot(600, capture)
    else:
        timer = QTimer(w)
        timer.timeout.connect(animate)
        timer.start(33)
    return app.exec()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
```

- [ ] **Paso 7: Correr los tests**

Correr: `venv/Scripts/python.exe -m pytest tests/test_main_window.py tests/test_widgets.py`
Esperado: `15 passed`.

- [ ] **Paso 8: Revisión visual (punto de control con el usuario)**

Correr:
```bash
venv/Scripts/python.exe -m knob.ui.preview Nebula --screenshot preview-nebula.png
venv/Scripts/python.exe -m knob.ui.preview Cobalt --screenshot preview-cobalt.png
```
Abrir los PNG con la herramienta Read y compararlos con la maqueta aprobada (`.superpowers/brainstorm/*/content/layout.html`, opción 3). Revisar que:
- La barra superior muestre logo, estado, perfil y ⚙.
- Haya 3 tarjetas de canal con ícono, nombre, fader con brillo, VU y %.
- Haya 2 tarjetas de encoder con perilla, título y ⚙.
- La barra de salida esté abajo.
- No haya texto cortado ni superpuesto.

Mostrarle las capturas al usuario y ajustar tamaños o espaciados si lo pide. Después borrar los PNG (no se commitean).

- [ ] **Paso 9: Commit**

```bash
git add knob/ui tests/test_main_window.py
git commit -m "feat: ventana principal con tarjetas de canal, encoders, salida de audio y vista previa" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 10: Superposiciones (asignar canal, encoder, perfiles y ajustes)

**Archivos:**
- Crear: `knob/ui/overlays/base.py`, `assign_channel.py`, `encoder_config.py`, `profile_name.py`, `settings_panel.py` (en `knob/ui/overlays/`)
- Modificar: `knob/ui/main_window.py` (crear `self.overlay`)
- Test: `tests/test_overlays.py`

**Interfaces:**
- Consume: theme, `IconBadge`, `Toggle`, `MainWindow.backdrop`.
- Produce `base`:
  - `OverlayHost(parent)` con señal `closed`, `open_centered(panel, width)`, `open_side(panel)`, `close_overlay()`, `is_open() -> bool` y `panel()`.
  - `Panel(parent=None, side=False)`, `section_label(text) -> QLabel`.
  - `OptionRow(text, theme, badge=None, selected=False, hint="")` con señal `clicked` (badge = tupla `(name, pixmap, kind)` o `None`).
- Produce `assign_channel`: `SYSTEM_OPTIONS`, `AssignChannelPanel(title, current, apps, icon_for, theme)` con señal `chosen(object)` (str o None).
- Produce `encoder_config`:
  - Listas `ENCODER_CLICK_ACTIONS` y `ENCODER_TURN_CHANNELS`, ambas `[(clave, etiqueta)]`.
  - `key_name(key: int, text: str) -> str|None` y `build_combo(ctrl, shift, alt, win, key, text) -> str|None`.
  - `HotkeyButton(combo)` con señal `changed(str)` y `combo()`.
  - `EncoderConfigPanel(enc_idx, turn_label, entry: dict, theme)` con señales `saved(dict)` y `cancelled()`.
- Produce `profile_name`: `ProfileNamePanel(title, initial, theme)` con señales `submitted(str)` y `cancelled()`, y `show_error(msg)`.
- Produce `settings_panel`: `SettingsPanel(theme)` con:
  - Señales `profile_switch(str)`, `profile_new()`, `profile_rename()`, `profile_delete()`, `theme_selected(str)`, `startup_toggled(bool)`, `tray_toggled(bool)`, `update_clicked()` y `close_requested()`.
  - Métodos `set_profiles(names, active)`, `set_theme(theme)`, `set_startup(bool)`, `set_tray(bool)`, `set_version(str)` y `set_update_state(state, text)` (state: `"none"`, `"available"`, `"downloading"` o `"error"`).

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_overlays.py`:
```python
# -*- coding: utf-8 -*-
from PySide6.QtCore import Qt


def theme():
    from knob.ui.theme import get_theme
    return get_theme("Nebula")


def test_build_combo():
    from knob.ui.overlays.encoder_config import build_combo, key_name
    K = Qt.Key
    assert build_combo(True, True, False, False, K.Key_M.value, "\x0d") == "ctrl+shift+m"
    assert build_combo(False, False, True, False, K.Key_F4.value, "") == "alt+f4"
    assert build_combo(False, False, False, True, K.Key_Space.value, " ") == "win+space"
    assert build_combo(True, False, False, False, K.Key_Control.value, "") is None
    assert key_name(K.Key_PageDown.value, "") == "pagedown"
    assert key_name(K.Key_5.value, "5") == "5"


def test_overlay_abre_y_cierra(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.overlays.base import Panel
    w = MainWindow(theme())
    w.show()
    closed = []
    w.overlay.closed.connect(lambda: closed.append(1))
    w.overlay.open_centered(Panel(), 300)
    assert w.overlay.is_open()
    w.overlay.open_side(Panel(side=True))         # reemplaza el panel anterior
    assert w.overlay.is_open()
    w.overlay.close_overlay()
    qapp.processEvents()
    w.overlay._finish_close()                     # sin esperar la animación
    assert not w.overlay.is_open()
    assert closed == [1]


def test_asignar_canal_emite_eleccion(qapp):
    from knob.ui.overlays.assign_channel import AssignChannelPanel
    from knob.ui.overlays.base import OptionRow
    got = []
    p = AssignChannelPanel("Canal 1", "Discord", ["Spotify"], lambda name: None, theme())
    p.chosen.connect(got.append)
    rows = p.findChildren(OptionRow)
    texts = [r.text for r in rows]
    assert texts == ["Volumen general", "Micrófono", "App en foco", "Discord", "Spotify"]
    assert [r.selected for r in rows] == [False, False, False, True, False]
    rows[4].clicked.emit()
    p.unassign_button.click()
    assert got == ["Spotify", None]
    assert not p.grab().isNull()


def test_asignar_sin_apps_muestra_mensaje(qapp):
    from knob.ui.overlays.assign_channel import AssignChannelPanel
    p = AssignChannelPanel("Canal 2", None, [], lambda name: None, theme())
    assert p.empty_label.isVisibleTo(p)
    assert p.unassign_button is None


def test_configurar_encoder_guarda(qapp):
    from knob.ui.overlays.encoder_config import EncoderConfigPanel
    entry = {"click": "mute", "dblclick": "hotkey", "click_turn": "output_device",
             "hotkey_keys_click": "", "hotkey_keys_dblclick": "ctrl+m"}
    p = EncoderConfigPanel(0, "Volumen general", entry, theme())
    assert p.hotkey_buttons["dblclick"].isVisibleTo(p)
    assert not p.hotkey_buttons["click"].isVisibleTo(p)
    saved = []
    p.saved.connect(saved.append)
    p.combos["click"].setCurrentIndex(2)              # Play / Pausa
    p.save_button.click()
    assert saved == [{"click": "play_pause", "dblclick": "hotkey", "click_turn": "output_device",
                      "hotkey_keys_click": "", "hotkey_keys_dblclick": "ctrl+m"}]


def test_nombre_de_perfil(qapp):
    from knob.ui.overlays.profile_name import ProfileNamePanel
    p = ProfileNamePanel("Nuevo perfil", "", theme())
    got = []
    p.submitted.connect(got.append)
    p.name_edit.setText("  Gaming ")
    p.save_button.click()
    assert got == ["Gaming"]
    p.show_error("Ya existe un perfil con ese nombre")
    assert p.error_label.text() == "Ya existe un perfil con ese nombre"


def test_ajustes(qapp):
    from knob.ui.overlays.base import OptionRow
    from knob.ui.overlays.settings_panel import SettingsPanel
    p = SettingsPanel(theme())
    p.set_profiles(["Default", "Gaming"], "Gaming")
    p.set_startup(True)
    p.set_tray(False)
    p.set_version("v4.0.0")
    p.set_update_state("available", "↑ v4.0.1 disponible")
    got = []
    p.profile_switch.connect(lambda n: got.append(("switch", n)))
    p.profile_delete.connect(lambda: got.append(("delete",)))
    p.theme_selected.connect(lambda n: got.append(("theme", n)))
    p.tray_toggled.connect(lambda b: got.append(("tray", b)))
    p.update_clicked.connect(lambda: got.append(("update",)))
    p.findChildren(OptionRow)[0].clicked.emit()
    p.btn_delete.click()                              # primer click: pide confirmación
    assert ("delete",) not in got
    p.btn_delete.click()                              # segundo click: borra
    p.swatches["Cobalt"].click()
    p.tg_tray.click()
    p.btn_update.click()
    assert got == [("switch", "Default"), ("delete",), ("theme", "Cobalt"), ("tray", True), ("update",)]
    p.set_profiles(["Default"], "Default")
    assert not p.btn_delete.isEnabled()
    assert not p.grab().isNull()
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_overlays.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.ui.overlays.encoder_config'`.

- [ ] **Paso 3: Implementar base**

`knob/ui/overlays/base.py`:
```python
# -*- coding: utf-8 -*-
"""Capa que oscurece la ventana y muestra paneles (centrados o deslizados desde la derecha)."""
from PySide6.QtCore import QEasingCurve, QParallelAnimationGroup, QPoint, QPropertyAnimation, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen
from PySide6.QtWidgets import QFrame, QGraphicsOpacityEffect, QHBoxLayout, QLabel, QWidget

from knob.ui.theme import PANEL_BG, qcolor, ui_font
from knob.ui.widgets.icon_badge import IconBadge

ANIM_MS = 190


def section_label(text: str) -> QLabel:
    lbl = QLabel(text.upper())
    lbl.setObjectName("section")
    f = ui_font(10, QFont.Weight.DemiBold)
    f.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.0)
    lbl.setFont(f)
    lbl.setContentsMargins(0, 10, 0, 2)
    return lbl


class Panel(QFrame):
    """Fondo de un panel: superficie oscura con borde suave."""

    def __init__(self, parent=None, side: bool = False):
        super().__init__(parent)
        self._side = side

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        if self._side:
            r.adjust(0, -14, 14, 14)            # solo se ven las esquinas izquierdas
        p.setPen(QPen(QColor(255, 255, 255, 23), 1))
        p.setBrush(PANEL_BG)
        p.drawRoundedRect(r, 12, 12)


class OptionRow(QWidget):
    """Fila clickeable de una lista: ícono opcional, texto, pista y tilde si está elegida."""
    clicked = Signal()

    def __init__(self, text: str, theme, badge=None, selected: bool = False, hint: str = "", parent=None):
        super().__init__(parent)
        self.text = text
        self.selected = selected
        self._t = theme
        self._hover = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(34)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(9, 5, 9, 5)
        lay.setSpacing(9)
        if badge is not None:
            b = IconBadge(20)
            b.set_theme(theme)
            b.set_content(*badge)
            lay.addWidget(b)
        label = QLabel(text)
        label.setFont(ui_font(12, QFont.Weight.DemiBold if selected else QFont.Weight.Normal))
        lay.addWidget(label, 1)
        if hint:
            h = QLabel(hint)
            h.setObjectName("dim")
            h.setFont(ui_font(10))
            lay.addWidget(h)
        if selected:
            lay.addWidget(QLabel("✓"))

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def enterEvent(self, e):
        self._hover = True
        self.update()

    def leaveEvent(self, e):
        self._hover = False
        self.update()

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton and self.rect().contains(e.position().toPoint()):
            self.clicked.emit()

    def mousePressEvent(self, e):
        e.accept()

    def paintEvent(self, _):
        if not (self.selected or self._hover):
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        if self.selected:
            g = QLinearGradient(0, 0, r.width(), 0)
            g.setColorAt(0, qcolor(self._t.grad_a, 115))
            g.setColorAt(1, qcolor(self._t.grad_b, 64))
            p.setBrush(g)
            p.setPen(QPen(qcolor(self._t.brand_a, 115), 1))
        else:
            p.setBrush(QColor(255, 255, 255, 14))
            p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(r, 8, 8)


class OverlayHost(QWidget):
    closed = Signal()

    def __init__(self, parent):
        super().__init__(parent)
        self._panel = None
        self._mode = None
        self._anim = None
        self._closing = False
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._effect = QGraphicsOpacityEffect(self)
        self._effect.setOpacity(1.0)
        self.setGraphicsEffect(self._effect)
        self.hide()

    def is_open(self) -> bool:
        return self._panel is not None and not self._closing

    def panel(self):
        return self._panel

    def open_centered(self, panel, width: int = 340) -> None:
        self._open(panel, "center", width)

    def open_side(self, panel) -> None:
        self._open(panel, "side", int(self.parentWidget().width() * 0.72))

    def _open(self, panel, mode: str, width: int) -> None:
        was_visible = self.isVisible() and not self._closing
        self._stop_anim()
        self._closing = False
        self._dispose_panel()
        self.setGeometry(self.parentWidget().rect())
        panel.setParent(self)
        self._panel, self._mode = panel, mode
        panel.setFixedWidth(width)
        if mode == "center":
            h = min(panel.sizeHint().height(), self.height() - 60)
            panel.setFixedHeight(h)
            end = QPoint((self.width() - width) // 2, max(30, (self.height() - h) // 3))
            start = end + QPoint(0, 14)
        else:
            panel.setFixedHeight(self.height())
            end = QPoint(self.width() - width, 0)
            start = QPoint(self.width(), 0)
        panel.move(start)
        panel.show()
        self.show()
        self.raise_()
        self.setFocus()

        group = QParallelAnimationGroup(self)
        slide = QPropertyAnimation(panel, b"pos", self)
        slide.setDuration(ANIM_MS)
        slide.setStartValue(start)
        slide.setEndValue(end)
        slide.setEasingCurve(QEasingCurve.Type.OutCubic)
        group.addAnimation(slide)
        if not was_visible:
            fade = QPropertyAnimation(self._effect, b"opacity", self)
            fade.setDuration(ANIM_MS)
            fade.setStartValue(0.0)
            fade.setEndValue(1.0)
            group.addAnimation(fade)
        else:
            self._effect.setOpacity(1.0)
        self._anim = group
        group.start()

    def close_overlay(self) -> None:
        if self._panel is None or self._closing:
            return
        self._stop_anim()
        self._closing = True
        fade = QPropertyAnimation(self._effect, b"opacity", self)
        fade.setDuration(150)
        fade.setStartValue(self._effect.opacity())
        fade.setEndValue(0.0)
        fade.finished.connect(self._finish_close)
        self._anim = fade
        fade.start()

    def _finish_close(self) -> None:
        if not self._closing:
            return
        self._closing = False
        self._stop_anim()
        self.hide()
        self._effect.setOpacity(1.0)
        self._dispose_panel()
        self.closed.emit()

    def _stop_anim(self) -> None:
        if self._anim is not None:
            try:
                self._anim.finished.disconnect()
            except (RuntimeError, TypeError):
                pass
            self._anim.stop()
            self._anim = None

    def _dispose_panel(self) -> None:
        if self._panel is not None:
            self._panel.hide()
            self._panel.deleteLater()
            self._panel = None

    def paintEvent(self, _):
        QPainter(self).fillRect(self.rect(), QColor(5, 3, 10, 140))

    def mousePressEvent(self, e):
        if self._panel is not None and not self._panel.geometry().contains(e.position().toPoint()):
            self.close_overlay()
        e.accept()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key.Key_Escape:
            self.close_overlay()
            return
        super().keyPressEvent(e)
```

- [ ] **Paso 4: Implementar assign_channel**

`knob/ui/overlays/assign_channel.py`:
```python
# -*- coding: utf-8 -*-
"""Panel 'Canal N · ¿Qué controla?': destinos de sistema + apps con audio."""
from PySide6.QtCore import Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget

from knob.ui.overlays.base import OptionRow, Panel, section_label
from knob.ui.theme import ui_font

SYSTEM_OPTIONS = (("master", "Volumen general", ""),
                  ("mic", "Micrófono", ""),
                  ("focus", "App en foco", "la ventana activa"))


class AssignChannelPanel(Panel):
    chosen = Signal(object)          # str con el destino, o None para quitar la asignación

    def __init__(self, title: str, current, apps: list, icon_for, theme, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 14, 14, 12)
        lay.setSpacing(2)
        head = QLabel(f"{title} · ¿Qué controla?")
        head.setFont(ui_font(14, QFont.Weight.Bold))
        lay.addWidget(head)

        lay.addWidget(section_label("Sistema"))
        for key, label, hint in SYSTEM_OPTIONS:
            row = OptionRow(label, theme, badge=("", None, key), selected=(current == key), hint=hint)
            row.clicked.connect(lambda k=key: self.chosen.emit(k))
            lay.addWidget(row)

        lay.addWidget(section_label("Apps con audio ahora"))
        names = list(apps)
        if current and current not in ("master", "mic", "focus") and current not in names:
            names.insert(0, current)          # la app asignada aunque no esté sonando
        inner = QWidget()
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(0, 0, 4, 0)
        inner_lay.setSpacing(2)
        for name in names:
            row = OptionRow(name, theme, badge=(name, icon_for(name), None), selected=(current == name))
            row.clicked.connect(lambda n=name: self.chosen.emit(n))
            inner_lay.addWidget(row)
        self.empty_label = QLabel("No hay apps reproduciendo audio.\nAbrí la app y volvé a intentar.")
        self.empty_label.setObjectName("dim")
        self.empty_label.setFont(ui_font(11))
        self.empty_label.setVisible(not names)
        inner_lay.addWidget(self.empty_label)
        inner_lay.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)
        scroll.setFixedHeight(min(220, max(48, 36 * len(names) + 8)))
        lay.addWidget(scroll)

        self.unassign_button = None
        if current is not None:
            self.unassign_button = QPushButton("Quitar asignación")
            self.unassign_button.setObjectName("link")
            self.unassign_button.clicked.connect(lambda: self.chosen.emit(None))
            lay.addWidget(self.unassign_button)
```

- [ ] **Paso 5: Implementar encoder_config**

`knob/ui/overlays/encoder_config.py`:
```python
# -*- coding: utf-8 -*-
"""Panel de acciones del botón de un encoder (click, doble click, mantener y girar)."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from knob.ui.overlays.base import Panel, section_label
from knob.ui.theme import ui_font

ENCODER_CLICK_ACTIONS = [
    ("none", "Ninguna"),
    ("mute", "Mute / Unmute"),
    ("play_pause", "Play / Pausa"),
    ("next_track", "Siguiente pista"),
    ("prev_track", "Pista anterior"),
    ("vol_up", "Vol. general +5%"),
    ("vol_down", "Vol. general -5%"),
    ("hotkey", "Atajo de teclado"),
]
ENCODER_TURN_CHANNELS = [
    ("none", "Canal asignado (normal)"),
    ("master", "Volumen general"),
    ("mic", "Micrófono"),
    ("focus", "App en foco"),
    ("output_device", "Cambiar salida de audio"),
]
ENCODER_NAMES = ("Encoder izquierdo", "Encoder derecho")

_K = Qt.Key
_SPECIAL = {
    _K.Key_Space.value: "space", _K.Key_Return.value: "enter", _K.Key_Enter.value: "enter",
    _K.Key_Escape.value: "esc", _K.Key_Tab.value: "tab", _K.Key_Backspace.value: "backspace",
    _K.Key_Delete.value: "delete", _K.Key_Insert.value: "insert", _K.Key_Home.value: "home",
    _K.Key_End.value: "end", _K.Key_PageUp.value: "pageup", _K.Key_PageDown.value: "pagedown",
    _K.Key_Up.value: "up", _K.Key_Down.value: "down", _K.Key_Left.value: "left", _K.Key_Right.value: "right",
}
MODIFIER_KEYS = {_K.Key_Control.value, _K.Key_Shift.value, _K.Key_Alt.value, _K.Key_Meta.value,
                 _K.Key_AltGr.value, _K.Key_Super_L.value, _K.Key_Super_R.value}


def _int(k) -> int:
    return k.value if hasattr(k, "value") else int(k)


def key_name(key: int, text: str):
    """Nombre de tecla en el formato que entiende system.send_hotkey (igual que v3.3)."""
    if key in _SPECIAL:
        return _SPECIAL[key]
    if _K.Key_F1.value <= key <= _K.Key_F12.value:
        return f"f{key - _K.Key_F1.value + 1}"
    if 0x20 < key < 0x7F:
        return chr(key).lower()
    if text and len(text) == 1 and text.isprintable():
        return text.lower()
    return None


def build_combo(ctrl: bool, shift: bool, alt: bool, win: bool, key: int, text: str):
    if key in MODIFIER_KEYS:
        return None
    name = key_name(key, text)
    if not name:
        return None
    mods = [m for m, on in (("ctrl", ctrl), ("shift", shift), ("alt", alt), ("win", win)) if on]
    return "+".join(mods + [name])


class HotkeyButton(QPushButton):
    changed = Signal(str)

    def __init__(self, combo: str, parent=None):
        super().__init__(parent)
        self._combo = combo or ""
        self._recording = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clicked.connect(self._start)
        self._refresh()

    def combo(self) -> str:
        return self._combo

    def _refresh(self) -> None:
        if self._recording:
            self.setText("Presioná las teclas…")
        else:
            self.setText(self._combo.replace("+", " + ").title() if self._combo else "Grabar atajo…")

    def _start(self) -> None:
        self._recording = True
        self._refresh()
        self.setFocus()
        self.grabKeyboard()

    def _stop(self) -> None:
        if self._recording:
            self._recording = False
            self.releaseKeyboard()
            self._refresh()

    def keyPressEvent(self, e):
        if not self._recording:
            super().keyPressEvent(e)
            return
        mods = e.modifiers()
        combo = build_combo(bool(mods & Qt.KeyboardModifier.ControlModifier),
                            bool(mods & Qt.KeyboardModifier.ShiftModifier),
                            bool(mods & Qt.KeyboardModifier.AltModifier),
                            bool(mods & Qt.KeyboardModifier.MetaModifier),
                            _int(e.key()), e.text())
        if combo:
            self._combo = combo
            self._stop()
            self.changed.emit(combo)

    def focusOutEvent(self, e):
        self._stop()
        super().focusOutEvent(e)


class EncoderConfigPanel(Panel):
    saved = Signal(dict)
    cancelled = Signal()

    def __init__(self, enc_idx: int, turn_label: str, entry: dict, theme, parent=None):
        super().__init__(parent)
        self.combos = {}
        self.hotkey_buttons = {}
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(4)
        head = QLabel(f"{ENCODER_NAMES[enc_idx]} · Botón")
        head.setFont(ui_font(14, QFont.Weight.Bold))
        lay.addWidget(head)
        sub = QLabel(f"Girar controla: <b>{turn_label}</b> (cambialo desde la tarjeta)")
        sub.setObjectName("dim")
        sub.setFont(ui_font(11))
        sub.setWordWrap(True)
        lay.addWidget(sub)

        for atype, label in (("click", "Click"), ("dblclick", "Doble click")):
            lay.addWidget(section_label(label))
            row = QHBoxLayout()
            row.setSpacing(6)
            combo = self._make_combo(ENCODER_CLICK_ACTIONS, entry.get(atype, "none"))
            row.addWidget(combo, 1)
            hk = HotkeyButton(entry.get(f"hotkey_keys_{atype}", ""))
            hk.setVisible(entry.get(atype) == "hotkey")
            row.addWidget(hk)
            combo.currentIndexChanged.connect(
                lambda _i, c=combo, h=hk: h.setVisible(c.currentData() == "hotkey"))
            lay.addLayout(row)
            self.combos[atype] = combo
            self.hotkey_buttons[atype] = hk

        lay.addWidget(section_label("Mantener y girar"))
        turn = self._make_combo(ENCODER_TURN_CHANNELS, entry.get("click_turn", "none"))
        lay.addWidget(turn)
        self.combos["click_turn"] = turn

        buttons = QHBoxLayout()
        buttons.setContentsMargins(0, 14, 0, 0)
        buttons.addStretch(1)
        cancel = QPushButton("Cancelar")
        cancel.clicked.connect(self.cancelled)
        buttons.addWidget(cancel)
        self.save_button = QPushButton("Guardar")
        self.save_button.setObjectName("primary")
        self.save_button.clicked.connect(self._save)
        buttons.addWidget(self.save_button)
        lay.addLayout(buttons)

    @staticmethod
    def _make_combo(options, current) -> QComboBox:
        combo = QComboBox()
        combo.setCursor(Qt.CursorShape.PointingHandCursor)
        for key, label in options:
            combo.addItem(label, key)
        idx = combo.findData(current)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        return combo

    def _save(self) -> None:
        self.saved.emit({
            "click": self.combos["click"].currentData(),
            "dblclick": self.combos["dblclick"].currentData(),
            "click_turn": self.combos["click_turn"].currentData(),
            "hotkey_keys_click": self.hotkey_buttons["click"].combo(),
            "hotkey_keys_dblclick": self.hotkey_buttons["dblclick"].combo(),
        })
```

- [ ] **Paso 6: Implementar profile_name y settings_panel**

`knob/ui/overlays/profile_name.py`:
```python
# -*- coding: utf-8 -*-
"""Panel para crear o renombrar un perfil."""
from PySide6.QtCore import Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QPushButton, QVBoxLayout

from knob.ui.overlays.base import Panel
from knob.ui.theme import ui_font


class ProfileNamePanel(Panel):
    submitted = Signal(str)
    cancelled = Signal()

    def __init__(self, title: str, initial: str, theme, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(8)
        head = QLabel(title)
        head.setFont(ui_font(14, QFont.Weight.Bold))
        lay.addWidget(head)
        self.name_edit = QLineEdit(initial)
        self.name_edit.setPlaceholderText("ej. Gaming, Trabajo…")
        self.name_edit.selectAll()
        self.name_edit.returnPressed.connect(self._submit)
        lay.addWidget(self.name_edit)
        self.error_label = QLabel("")
        self.error_label.setObjectName("error")
        self.error_label.setFont(ui_font(11))
        self.error_label.hide()
        lay.addWidget(self.error_label)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton("Cancelar")
        cancel.clicked.connect(self.cancelled)
        buttons.addWidget(cancel)
        self.save_button = QPushButton("Guardar")
        self.save_button.setObjectName("primary")
        self.save_button.clicked.connect(self._submit)
        buttons.addWidget(self.save_button)
        lay.addLayout(buttons)

    def showEvent(self, e):
        super().showEvent(e)
        self.name_edit.setFocus()

    def _submit(self) -> None:
        name = self.name_edit.text().strip()
        if not name:
            self.show_error("El nombre no puede estar vacío")
            return
        self.submitted.emit(name)

    def show_error(self, msg: str) -> None:
        self.error_label.setText(msg)
        self.error_label.show()
```

`knob/ui/overlays/settings_panel.py`:
```python
# -*- coding: utf-8 -*-
"""Panel de ajustes deslizable: perfiles, tema, sistema y versión."""
from PySide6.QtCore import QPointF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen
from PySide6.QtWidgets import (QAbstractButton, QHBoxLayout, QLabel, QPushButton, QScrollArea,
                               QToolButton, QVBoxLayout, QWidget)

from knob.ui.overlays.base import OptionRow, Panel, section_label
from knob.ui.theme import GLYPH_CLOSE, THEME_ORDER, THEMES, icon_font, ui_font
from knob.ui.widgets.toggle import Toggle

DELETE_CONFIRM_MS = 4000


class SwatchButton(QAbstractButton):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._swatch = theme
        self.selected = False
        self.setFixedSize(30, 30)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(theme.name)

    def sizeHint(self):
        return QSize(30, 30)

    def set_selected(self, on: bool) -> None:
        self.selected = on
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        c = QPointF(self.width() / 2, self.height() / 2)
        g = QLinearGradient(0, 0, self.width(), self.height())
        g.setColorAt(0, QColor(self._swatch.grad_a))
        g.setColorAt(1, QColor(self._swatch.grad_b))
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(g)
        p.drawEllipse(c, 11, 11)
        if self.selected:
            p.setPen(QPen(QColor("#FFFFFF"), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawEllipse(c, 14, 14)


class SettingsPanel(Panel):
    profile_switch = Signal(str)
    profile_new = Signal()
    profile_rename = Signal()
    profile_delete = Signal()
    theme_selected = Signal(str)
    startup_toggled = Signal(bool)
    tray_toggled = Signal(bool)
    update_clicked = Signal()
    close_requested = Signal()

    def __init__(self, theme, parent=None):
        super().__init__(parent, side=True)
        self._t = theme
        self._active = ""
        self._confirm_delete = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(4)

        head = QHBoxLayout()
        title = QLabel("Ajustes")
        title.setFont(ui_font(16, QFont.Weight.Bold))
        head.addWidget(title)
        head.addStretch(1)
        close = QToolButton()
        close.setFont(icon_font(10))
        close.setText(GLYPH_CLOSE)
        close.setFixedSize(26, 26)
        close.setCursor(Qt.CursorShape.PointingHandCursor)
        close.clicked.connect(self.close_requested)
        head.addWidget(close)
        lay.addLayout(head)

        lay.addWidget(section_label("Perfiles"))
        self._profiles_widget = QWidget()
        self._profiles_layout = QVBoxLayout(self._profiles_widget)
        self._profiles_layout.setContentsMargins(0, 0, 4, 0)
        self._profiles_layout.setSpacing(2)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._profiles_widget)
        scroll.setMaximumHeight(150)
        self._profiles_scroll = scroll
        lay.addWidget(scroll)
        row = QHBoxLayout()
        row.setSpacing(6)
        self.btn_new = QPushButton("+ Nuevo")
        self.btn_rename = QPushButton("Renombrar")
        self.btn_delete = QPushButton("Borrar")
        self.btn_delete.setObjectName("danger")
        self.btn_new.clicked.connect(self.profile_new)
        self.btn_rename.clicked.connect(self.profile_rename)
        self.btn_delete.clicked.connect(self._on_delete)
        for b in (self.btn_new, self.btn_rename, self.btn_delete):
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            row.addWidget(b)
        row.addStretch(1)
        lay.addLayout(row)

        lay.addWidget(section_label("Tema"))
        sw_row = QHBoxLayout()
        sw_row.setSpacing(6)
        self.swatches = {}
        for name in THEME_ORDER:
            b = SwatchButton(THEMES[name])
            b.clicked.connect(lambda _=False, n=name: self.theme_selected.emit(n))
            sw_row.addWidget(b)
            self.swatches[name] = b
        sw_row.addStretch(1)
        lay.addLayout(sw_row)

        lay.addWidget(section_label("Sistema"))
        self.tg_startup = Toggle(theme)
        self.tg_tray = Toggle(theme)
        self.tg_startup.toggled.connect(self.startup_toggled)
        self.tg_tray.toggled.connect(self.tray_toggled)
        lay.addLayout(self._setting_row("Iniciar con Windows", self.tg_startup))
        lay.addLayout(self._setting_row("Al cerrar, ir a la bandeja", self.tg_tray))

        lay.addWidget(section_label("Acerca de"))
        about = QHBoxLayout()
        self.lbl_version = QLabel("")
        self.lbl_version.setFont(ui_font(12))
        about.addWidget(self.lbl_version)
        about.addStretch(1)
        self.btn_update = QPushButton("")
        self.btn_update.setObjectName("primary")
        self.btn_update.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_update.clicked.connect(self.update_clicked)
        self.btn_update.hide()
        about.addWidget(self.btn_update)
        lay.addLayout(about)
        lay.addStretch(1)
        self.set_theme(theme)

    @staticmethod
    def _setting_row(text: str, toggle) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setContentsMargins(0, 4, 0, 4)
        lbl = QLabel(text)
        lbl.setFont(ui_font(12))
        row.addWidget(lbl)
        row.addStretch(1)
        row.addWidget(toggle)
        return row

    # --- API ---------------------------------------------------------------
    def set_profiles(self, names, active: str) -> None:
        self._active = active
        while self._profiles_layout.count():
            item = self._profiles_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for name in names:
            r = OptionRow(name, self._t, selected=(name == active))
            r.clicked.connect(lambda n=name: self.profile_switch.emit(n))
            self._profiles_layout.addWidget(r)
        self._profiles_scroll.setFixedHeight(min(150, 36 * len(names) + 4))
        self.btn_delete.setEnabled(len(names) > 1)
        self._reset_delete()

    def set_theme(self, theme) -> None:
        self._t = theme
        self.tg_startup.set_theme(theme)
        self.tg_tray.set_theme(theme)
        for name, b in self.swatches.items():
            b.set_selected(name == theme.name)
        for r in self._profiles_widget.findChildren(OptionRow):
            r.set_theme(theme)

    def set_startup(self, on: bool) -> None:
        self.tg_startup.blockSignals(True)
        self.tg_startup.setChecked(on)
        self.tg_startup.blockSignals(False)

    def set_tray(self, on: bool) -> None:
        self.tg_tray.blockSignals(True)
        self.tg_tray.setChecked(on)
        self.tg_tray.blockSignals(False)

    def set_version(self, version: str) -> None:
        self.lbl_version.setText(f"KNOB {version}")

    def set_update_state(self, state: str, text: str) -> None:
        self.btn_update.setVisible(state != "none")
        self.btn_update.setText(text)
        self.btn_update.setEnabled(state == "available")

    # --- borrar con confirmación -----------------------------------------------
    def _on_delete(self) -> None:
        if not self._confirm_delete:
            self._confirm_delete = True
            self.btn_delete.setText(f"¿Borrar «{self._active}»?")
            QTimer.singleShot(DELETE_CONFIRM_MS, self._reset_delete)
            return
        self._reset_delete()
        self.profile_delete.emit()

    def _reset_delete(self) -> None:
        self._confirm_delete = False
        self.btn_delete.setText("Borrar")
```

- [ ] **Paso 7: Conectar el host a la ventana**

En `knob/ui/main_window.py`, agregar el import:
```python
from knob.ui.overlays.base import OverlayHost
```
Y al final de `MainWindow.__init__` (después de `root.addWidget(self.output_bar)`):
```python
        self.overlay = OverlayHost(self.backdrop)
```

- [ ] **Paso 8: Correr y ver que pasa**

Correr: `venv/Scripts/python.exe -m pytest`
Esperado: todos los tests pasan (`7 passed` en `test_overlays.py`; ningún fallo en el resto).

- [ ] **Paso 9: Vista previa de las superposiciones**

En `knob/ui/preview.py`, agregar este soporte de `--overlay NOMBRE` (`assign`, `encoder` o `settings`). Va al principio del archivo, debajo de los imports:
```python
from knob.ui.overlays.assign_channel import AssignChannelPanel
from knob.ui.overlays.encoder_config import EncoderConfigPanel
from knob.ui.overlays.settings_panel import SettingsPanel


def open_demo_overlay(w, theme, which):
    if which == "assign":
        w.overlay.open_centered(AssignChannelPanel(
            "Canal 1", "Opera", ["Opera", "Discord", "Spotify", "Chrome"], lambda n: None, theme), 340)
    elif which == "encoder":
        w.overlay.open_centered(EncoderConfigPanel(0, "Volumen general", {
            "click": "play_pause", "dblclick": "hotkey", "click_turn": "output_device",
            "hotkey_keys_dblclick": "ctrl+shift+m"}, theme), 360)
    elif which == "settings":
        p = SettingsPanel(theme)
        p.set_profiles(["Default", "Gaming", "Trabajo"], "Gaming")
        p.set_startup(True)
        p.set_version("v4.0.0")
        p.set_update_state("available", "↑ v4.0.1 disponible")
        w.overlay.open_side(p)
```
En `main()`:
- Después de `shot = ...`, agregar:
```python
    overlay = argv[argv.index("--overlay") + 1] if "--overlay" in argv else None
    if overlay in args:
        args.remove(overlay)
```
- Después de `w.show()`, agregar:
```python
    if overlay:
        QTimer.singleShot(50, lambda: open_demo_overlay(w, theme, overlay))
```

Correr las tres capturas y revisarlas con Read:
```bash
for o in assign encoder settings; do venv/Scripts/python.exe -m knob.ui.preview Nebula --overlay $o --screenshot overlay-$o.png; done
```
Esperado: cada panel se ve completo, sin cortes, con el fondo oscurecido y el estilo de la maqueta aprobada (`interacciones-v2.html`). Mostrarle las capturas al usuario y después borrar los PNG.

- [ ] **Paso 10: Commit**

```bash
git add knob/ui tests/test_overlays.py
git commit -m "feat: superposiciones para asignar canal, configurar encoder, perfiles y ajustes" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 11: Bandeja, controlador y punto de entrada

**Archivos:**
- Crear: `knob/ui/tray.py`, `knob/controller.py`, `main.py`
- Test: `tests/test_imports.py` + prueba real con el KNOB

**Interfaces:**
- Consume: todo lo anterior.
- Produce:
  - `Tray(icon: QIcon)` con señales `show_requested` y `quit_requested`, métodos `show()` y `hide()`, y el estático `available() -> bool`.
  - `Controller(app: QApplication, start_minimized: bool)` con `start()`, `quit()`, `hide_to_tray()` y `show_window()`. Implementa el protocolo `Effects` del motor.
  - `main.py` con `main() -> int`.

- [ ] **Paso 1: Escribir el test que falla**

`tests/test_imports.py`:
```python
# -*- coding: utf-8 -*-


def test_modulos_importan(qapp):
    import knob.controller  # noqa: F401
    import knob.ui.tray  # noqa: F401
    import main  # noqa: F401


def test_tray_se_construye(qapp):
    from PySide6.QtGui import QIcon
    from knob.core.paths import resource_path
    from knob.ui.tray import Tray
    t = Tray(QIcon(resource_path("icon.png")))
    got = []
    t.show_requested.connect(lambda: got.append("show"))
    t._menu.actions()[0].trigger()
    assert got == ["show"]
```

- [ ] **Paso 2: Correr y ver que falla**

Correr: `venv/Scripts/python.exe -m pytest tests/test_imports.py`
Esperado: FAIL con `ModuleNotFoundError: No module named 'knob.controller'`.

- [ ] **Paso 3: Implementar tray**

`knob/ui/tray.py`:
```python
# -*- coding: utf-8 -*-
"""Ícono en la bandeja del sistema (se muestra solo mientras la ventana está oculta, como v3.3)."""
from PySide6.QtCore import QObject, Signal
from PySide6.QtWidgets import QMenu, QSystemTrayIcon


class Tray(QObject):
    show_requested = Signal()
    quit_requested = Signal()

    def __init__(self, icon, parent=None):
        super().__init__(parent)
        self._tray = QSystemTrayIcon(icon)
        self._tray.setToolTip("KNOB")
        self._menu = QMenu()
        self._menu.addAction("Mostrar KNOB").triggered.connect(lambda: self.show_requested.emit())
        self._menu.addAction("Salir").triggered.connect(lambda: self.quit_requested.emit())
        self._tray.setContextMenu(self._menu)
        self._tray.activated.connect(self._on_activated)

    @staticmethod
    def available() -> bool:
        return QSystemTrayIcon.isSystemTrayAvailable()

    def _on_activated(self, reason) -> None:
        if reason in (QSystemTrayIcon.ActivationReason.Trigger,
                      QSystemTrayIcon.ActivationReason.DoubleClick):
            self.show_requested.emit()

    def show(self) -> None:
        self._tray.show()

    def hide(self) -> None:
        self._tray.hide()
```

- [ ] **Paso 4: Implementar el controlador**

`knob/controller.py`:
```python
# -*- coding: utf-8 -*-
"""Une el motor, el hardware, el audio de Windows y la interfaz."""
import threading

from PySide6.QtCore import QObject, QTimer, Signal
from PySide6.QtGui import QIcon

from knob.core import system, updater
from knob.core.audio import AudioBackend, OutputDevices, find_exe_path
from knob.core.config import Config, ProfileError
from knob.core.engine import Engine, linear_to_vu, next_peak
from knob.core.log import log
from knob.core.paths import resource_path
from knob.core.serial_link import SerialLink
from knob.ui.icons import AppIcons
from knob.ui.main_window import MainWindow
from knob.ui.overlays.assign_channel import AssignChannelPanel
from knob.ui.overlays.encoder_config import EncoderConfigPanel
from knob.ui.overlays.profile_name import ProfileNamePanel
from knob.ui.overlays.settings_panel import SettingsPanel
from knob.ui.theme import build_qss, get_theme
from knob.ui.tray import Tray

SYSTEM_LABELS = {"master": "Volumen general", "mic": "Micrófono", "focus": "App en foco"}
CHANNEL_TITLES = ("Canal 1", "Canal 2", "Canal 3", "Encoder izquierdo", "Encoder derecho")
STATUS_TEXT = {"connected": "Conectado", "searching": "Buscando KNOB…"}


class Controller(QObject):
    _output_loaded = Signal(str)
    _update_available = Signal(str, str)
    _update_progress = Signal(int)
    _update_failed = Signal()
    _update_ready = Signal(str)

    def __init__(self, app, start_minimized: bool = False):
        super().__init__()
        self.app = app
        self._start_minimized = start_minimized
        self.config = Config()
        self.theme = get_theme(self.config.theme)
        self.audio = AudioBackend()
        self.outputs = OutputDevices()
        self.focus = system.FocusTracker()
        self.engine = Engine(effects=self, focus_provider=lambda: self.focus.current)
        self.engine.set_assignments(self.config.channel_assignments(), self.config.encoder_assignments())
        self.icons = AppIcons(self._exe_lookup)
        self.window = MainWindow(self.theme)
        self.serial = SerialLink(lambda: self.config.last_port)
        self.tray = Tray(QIcon(resource_path("icon.png")))
        self._version = updater.read_version()
        self._update = None
        self._update_state = ("none", "")
        self._settings = None
        self._vu_peak = [0.0, 0.0, 0.0]
        self._connect_signals()
        self._timers = []
        for interval, slot in ((16, self._tick), (33, self._update_vu),
                               (300, self._update_focus_labels), (10000, self._refresh_cards)):
            timer = QTimer(self)
            timer.setInterval(interval)
            timer.timeout.connect(slot)
            self._timers.append(timer)

    # --- arranque / cierre ---------------------------------------------------
    def _connect_signals(self) -> None:
        w = self.window
        w.channel_clicked.connect(self._open_assign)
        w.encoder_gear_clicked.connect(self._open_encoder)
        w.settings_clicked.connect(self._open_settings)
        w.profile_selected.connect(self._switch_profile)
        w.output_prev.connect(lambda: self.cycle_output(-1))
        w.output_next.connect(lambda: self.cycle_output(1))
        w.close_requested.connect(self._on_close_requested)
        w.overlay.closed.connect(self._on_overlay_closed)
        # Señales que vienen de hilos: siempre a métodos del controlador (QObject del hilo principal)
        # para que Qt las encole y se ejecuten en el hilo de la interfaz.
        self.serial.button_event.connect(self._on_button)
        self.serial.status_changed.connect(self._on_status)
        self.serial.port_connected.connect(self._on_port_connected)
        self.tray.show_requested.connect(self.show_window)
        self.tray.quit_requested.connect(self.quit)
        self._output_loaded.connect(self._on_output_loaded)
        self._update_available.connect(self._on_update_available)
        self._update_progress.connect(self._on_update_progress)
        self._update_failed.connect(self._on_update_failed)
        self._update_ready.connect(self._on_update_ready)

    def start(self) -> None:
        self.app.setStyleSheet(build_qss(self.theme))
        self.window.apply_theme(self.theme)
        self._refresh_cards()
        self._refresh_profiles()
        self.window.status.set_state("searching", STATUS_TEXT["searching"])
        system.refresh_startup_path()
        self.audio.start()
        self.focus.start()
        self.serial.start()
        threading.Thread(target=lambda: self._output_loaded.emit(self.outputs.load() or ""), daemon=True).start()
        threading.Thread(target=self._check_update, daemon=True).start()
        for timer in self._timers:
            timer.start()
        if self._start_minimized and Tray.available():
            self.hide_to_tray()
        else:
            self.window.show()

    def _on_close_requested(self) -> None:
        if self.config.minimize_to_tray and Tray.available():
            self.hide_to_tray()
        else:
            self.quit()

    def hide_to_tray(self) -> None:
        self.window.overlay.close_overlay()
        self.window.hide()
        self.tray.show()

    def show_window(self) -> None:
        self.tray.hide()
        self.window.showNormal()
        self.window.raise_()
        self.window.activateWindow()

    def quit(self) -> None:
        for timer in self._timers:
            timer.stop()
        self.serial.stop()
        self.tray.hide()
        self.app.quit()

    # --- Effects (los llama el motor) ---------------------------------------------
    def set_volume(self, target: str, level: float) -> None:
        try:
            self.audio.set_volume(target, level)
        except Exception:
            pass        # sesión muerta o dispositivo desconectado: se reintenta en el próximo tick

    def media_key(self, action: str) -> None:
        system.send_media_key(action)

    def send_hotkey(self, combo: str) -> None:
        threading.Thread(target=system.send_hotkey, args=(combo,), daemon=True).start()

    def step_master(self, delta: float) -> None:
        try:
            self.audio.step_master(delta)
        except Exception as e:
            log(f"step_master: {e}")

    def cycle_output(self, direction: int) -> None:
        name = self.outputs.cycle(direction)
        if name:
            self.window.output_bar.set_name(name, flash=True)

    # --- bucles --------------------------------------------------------------
    def _tick(self) -> None:
        vals = self.serial.take_latest()
        if vals is not None:
            self.engine.on_frame(vals)
        self.engine.tick()
        for ch in range(5):
            card = self.window.card_for(ch)
            card.set_value(self.engine.smooth_vals[ch])
            card.set_muted(self.engine.is_muted(ch))

    def _update_vu(self) -> None:
        for ch in range(3):
            assignment = self.engine.channel_assignments.get(ch)
            raw = self.audio.peak(assignment, self.focus.current) if assignment else 0.0
            raw *= self.engine.smooth_vals[ch] / 1023.0      # escala con la posición del fader
            level = linear_to_vu(raw)
            self._vu_peak[ch] = next_peak(level, self._vu_peak[ch])
            self.window.channel_cards[ch].set_level(level, self._vu_peak[ch])

    def _update_focus_labels(self) -> None:
        for ch in range(5):
            if self.engine.channel_assignments.get(ch) == "focus":
                self.window.card_for(ch).set_focus_target(self.focus.current or None)

    # --- tarjetas ------------------------------------------------------------
    def _exe_lookup(self, app_name: str):
        path = find_exe_path(app_name)
        if path:
            self.config.remember_exe_path(app_name, path)
            return path
        return self.config.exe_path(app_name)

    def _describe(self, assignment):
        if assignment is None:
            return None, None, None
        if assignment in SYSTEM_LABELS:
            return SYSTEM_LABELS[assignment], None, assignment
        return assignment, self.icons.icon_for(assignment), None

    def _refresh_cards(self) -> None:
        for ch in range(5):
            self.window.card_for(ch).set_assignment(*self._describe(self.engine.channel_assignments.get(ch)))
        self._update_focus_labels()

    def _on_status(self, state: str, detail: str) -> None:
        if state == "busy":
            text = f"{detail} ocupado: cerrá el Serial Monitor"
        else:
            text = STATUS_TEXT.get(state, state)
        self.window.status.set_state(state, text)

    def _on_port_connected(self, port: str) -> None:
        self.config.last_port = port

    def _on_button(self, enc_idx: int, code: int) -> None:
        self.engine.on_button(enc_idx, code)

    def _on_output_loaded(self, name: str) -> None:
        self.window.output_bar.set_name(name or None)

    # --- asignar canal / encoder ---------------------------------------------
    def _open_assign(self, ch: int) -> None:
        self._settings = None
        panel = AssignChannelPanel(CHANNEL_TITLES[ch], self.engine.channel_assignments.get(ch),
                                   self.audio.app_names(), self.icons.icon_for, self.theme)
        panel.chosen.connect(lambda value, ch=ch: self._assign(ch, value))
        self.window.overlay.open_centered(panel, 340)

    def _assign(self, ch: int, value) -> None:
        self.engine.channel_assignments[ch] = value
        self.config.set_channel_assignment(ch, value)
        self._refresh_cards()
        self.window.overlay.close_overlay()

    def _open_encoder(self, enc: int) -> None:
        self._settings = None
        title = self._describe(self.engine.channel_assignments.get(3 + enc))[0] or "Sin asignar"
        panel = EncoderConfigPanel(enc, title, self.engine.encoder_assignments[enc], self.theme)
        panel.saved.connect(lambda entry, enc=enc: self._save_encoder(enc, entry))
        panel.cancelled.connect(self.window.overlay.close_overlay)
        self.window.overlay.open_centered(panel, 360)

    def _save_encoder(self, enc: int, entry: dict) -> None:
        self.engine.encoder_assignments[enc] = dict(entry)
        self.config.set_encoder_assignments(self.engine.encoder_assignments)
        self.window.overlay.close_overlay()

    # --- ajustes y perfiles ----------------------------------------------------------
    def _open_settings(self) -> None:
        p = SettingsPanel(self.theme)
        p.set_profiles(self.config.profile_names(), self.config.active_profile)
        p.set_startup(system.is_startup_enabled())
        p.set_tray(self.config.minimize_to_tray)
        p.set_version(self._version)
        p.set_update_state(*self._update_state)
        p.profile_switch.connect(self._switch_profile)
        p.profile_new.connect(lambda: self._open_profile_name("new"))
        p.profile_rename.connect(lambda: self._open_profile_name("rename"))
        p.profile_delete.connect(self._delete_profile)
        p.theme_selected.connect(self._set_theme)
        p.startup_toggled.connect(system.set_startup_enabled)
        p.tray_toggled.connect(self._set_tray)
        p.update_clicked.connect(self._start_update)
        p.close_requested.connect(self.window.overlay.close_overlay)
        self._settings = p
        self.window.overlay.open_side(p)

    def _on_overlay_closed(self) -> None:
        self._settings = None

    def _open_profile_name(self, mode: str) -> None:
        active = self.config.active_profile
        panel = ProfileNamePanel("Nuevo perfil" if mode == "new" else "Renombrar perfil",
                                 "" if mode == "new" else active, self.theme)

        def submit(name):
            try:
                if mode == "new":
                    self.config.create_profile(name, self.engine.channel_assignments,
                                               self.engine.encoder_assignments)
                else:
                    self.config.rename_profile(active, name)
            except ProfileError as e:
                panel.show_error(str(e))
                return
            self._refresh_profiles()
            self._open_settings()

        panel.submitted.connect(submit)
        panel.cancelled.connect(self._open_settings)
        self._settings = None
        self.window.overlay.open_centered(panel, 320)

    def _switch_profile(self, name: str) -> None:
        try:
            self.config.switch_profile(name)
        except ProfileError:
            return
        self._apply_active_profile()

    def _delete_profile(self) -> None:
        if self.config.delete_profile(self.config.active_profile):
            self._apply_active_profile()

    def _apply_active_profile(self) -> None:
        self.engine.set_assignments(self.config.channel_assignments(), self.config.encoder_assignments())
        self._refresh_cards()
        self._refresh_profiles()

    def _refresh_profiles(self) -> None:
        names, active = self.config.profile_names(), self.config.active_profile
        self.window.set_profiles(names, active)
        if self._settings is not None:
            self._settings.set_profiles(names, active)

    def _set_theme(self, name: str) -> None:
        self.config.theme = name
        self.theme = get_theme(name)
        self.app.setStyleSheet(build_qss(self.theme))
        self.window.apply_theme(self.theme)
        if self._settings is not None:
            self._settings.set_theme(self.theme)

    def _set_tray(self, on: bool) -> None:
        self.config.minimize_to_tray = on

    # --- actualizaciones --------------------------------------------------------------
    def _check_update(self) -> None:
        result = updater.check_for_update(self._version)
        if result:
            self._update_available.emit(*result)

    def _on_update_available(self, tag: str, url: str) -> None:
        self._update = (tag, url)
        self._set_update_state("available", f"↑ {tag} disponible")
        self.window.set_update_badge(True)

    def _on_update_progress(self, pct: int) -> None:
        self._set_update_state("downloading", f"Descargando {pct}%")

    def _on_update_failed(self) -> None:
        self._set_update_state("error", "Error al descargar")

    def _set_update_state(self, state: str, text: str) -> None:
        self._update_state = (state, text)
        if self._settings is not None:
            self._settings.set_update_state(state, text)

    def _start_update(self) -> None:
        if not self._update or self._update_state[0] != "available":
            return
        url = self._update[1]
        self._set_update_state("downloading", "Descargando 0%")

        def run():
            try:
                path = updater.download(url, self._update_progress.emit)
                self._update_ready.emit(path)
            except Exception as e:
                log(f"descarga de actualización: {e}")
                self._update_failed.emit()

        threading.Thread(target=run, daemon=True).start()

    def _on_update_ready(self, path: str) -> None:
        updater.launch_installer(path)
        QTimer.singleShot(500, self.quit)     # salir de verdad para que el instalador pueda reemplazar archivos
```

- [ ] **Paso 5: Implementar main.py**

`main.py`:
```python
# -*- coding: utf-8 -*-
"""KNOB App: punto de entrada."""
import sys

from knob.core import system


def main() -> int:
    if not system.acquire_single_instance():
        system.show_already_running()
        return 0
    system.set_app_user_model_id("KNOB.App")
    import comtypes  # noqa: F401  inicializa COM (STA) en el hilo principal antes que Qt
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication

    from knob.controller import Controller
    from knob.core.paths import resource_path
    from knob.ui.theme import ui_font

    app = QApplication(sys.argv)
    app.setApplicationName("KNOB")
    app.setStyle("Fusion")
    app.setFont(ui_font(13))
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(QIcon(resource_path("icon.png")))
    controller = Controller(app, start_minimized="--minimized" in sys.argv)
    controller.start()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Paso 6: Correr todos los tests**

Correr: `venv/Scripts/python.exe -m pytest`
Esperado: todos pasan, incluidos los 2 de `test_imports.py`.

- [ ] **Paso 7: Prueba real con el KNOB**

La v3.3 instalada no deja abrir una segunda instancia (mutex). Pedirle al usuario que la cierre desde la bandeja (clic derecho → Salir). Después:
```bash
venv/Scripts/python.exe main.py
```
Verificar con el usuario y con `knob_debug.log`:
1. La pastilla pasa a "Conectado".
2. Los faders mueven las tarjetas y el volumen de cada app en el mezclador de Windows.
3. Los VU se mueven con el audio.
4. Al hacer click en un canal aparece la lista con íconos reales y se puede asignar.
5. Los ajustes abren el panel lateral y el cambio de tema funciona en vivo.
6. Cerrar la ventana respeta "Al cerrar, ir a la bandeja".

Arreglar lo que falle antes de seguir.

- [ ] **Paso 8: Commit**

```bash
git add knob/ui/tray.py knob/controller.py main.py tests/test_imports.py
git commit -m "feat: controlador que une motor, serial, audio e interfaz; bandeja y punto de entrada" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 12: Empaquetado v4.0.0 (PyInstaller en carpeta + Inno Setup)

**Archivos:**
- Modificar: `KNOB.spec` (reemplazo completo), `KNOB_App_Installer.iss` (reemplazo completo), `release.py` (paso 1: actualizar también `AppVersion` del .iss)

**Interfaces:**
- Consume: `main.py`, `icon.png`, `icon.ico`, `fonts/`, `version.txt`.
- Produce: `dist\KNOB\KNOB.exe` + `dist\KNOB\_internal\`, y `Output\KNOB_App_Installer.exe`.

- [ ] **Paso 1: Reemplazar `KNOB.spec`**

```python
# -*- mode: python ; coding: utf-8 -*-
# KNOB v4: build en carpeta (onedir) para que arranque al instante.

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.png', '.'), ('fonts', 'fonts'), ('version.txt', '.')],
    hiddenimports=['pycaw.pycaw', 'comtypes', 'comtypes.client', 'comtypes.server'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 'customtkinter', 'pystray', 'PIL',
        'PySide6.QtNetwork', 'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtQuickWidgets',
        'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtMultimedia',
        'PySide6.QtPdf', 'PySide6.QtSql', 'PySide6.QtTest', 'PySide6.QtOpenGL',
        'PySide6.QtOpenGLWidgets', 'PySide6.QtDBus', 'PySide6.QtXml', 'PySide6.QtConcurrent',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='KNOB',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='KNOB',
)
```

- [ ] **Paso 2: Reemplazar `KNOB_App_Installer.iss`**

Mantener `AppName` y `DefaultDirName` iguales que v3.3, para que actualice sobre la instalación existente.
```ini
; Script de Inno Setup para KNOB App
[Setup]
AppName=KNOB App
AppVersion=4.0.0
DefaultDirName={pf}\KNOB App
DefaultGroupName=KNOB App
OutputBaseFilename=KNOB_App_Installer
Compression=lzma
SolidCompression=yes
SetupIconFile=E:\KNOB\icon.ico
UninstallDisplayIcon={app}\KNOB.exe
CloseApplications=force

[Files]
Source: "E:\KNOB\dist\KNOB\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "E:\KNOB\version.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\KNOB App"; Filename: "{app}\KNOB.exe"
Name: "{group}\Desinstalar KNOB App"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\KNOB.exe"; Description: "Iniciar KNOB App"; Flags: nowait postinstall skipifsilent
```

- [ ] **Paso 3: `release.py` actualiza también el .iss**

En `release.py`, reemplazar el bloque del paso 1:
```python
# 1. Actualizar version.txt
print(f"[1/4] Actualizando version.txt -> {VERSION}")
with open(os.path.join(ROOT, "version.txt"), "w", encoding="utf-8") as f:
    f.write(VERSION + "\n")
```
por:
```python
# 1. Actualizar version.txt y AppVersion del instalador
print(f"[1/4] Actualizando version.txt e instalador -> {VERSION}")
with open(os.path.join(ROOT, "version.txt"), "w", encoding="utf-8") as f:
    f.write(VERSION + "\n")
iss_path = os.path.join(ROOT, "KNOB_App_Installer.iss")
with open(iss_path, encoding="utf-8") as f:
    iss = re.sub(r"(?m)^AppVersion=.*$", f"AppVersion={VERSION.lstrip('v')}", f.read())
with open(iss_path, "w", encoding="utf-8") as f:
    f.write(iss)
```
y cambiar la línea `import sys, os, subprocess, json, urllib.request` por `import sys, os, re, subprocess, json, urllib.request`.

- [ ] **Paso 4: Build local (sin publicar)**

```bash
printf 'v4.0.0\n' > version.txt
venv/Scripts/pyinstaller.exe --clean -y KNOB.spec
ls dist/KNOB/KNOB.exe dist/KNOB/_internal/version.txt dist/KNOB/_internal/fonts/AmazingViews.ttf
du -sh dist/KNOB
```
Esperado: los tres archivos existen y la carpeta pesa menos de ~120 MB. Si pesa mucho más, revisar con `ls dist/KNOB/_internal/PySide6` qué módulos Qt se colaron y agregarlos a `excludes`.

Después compilar el instalador:
```bash
"/c/Program Files (x86)/Inno Setup 6/ISCC.exe" KNOB_App_Installer.iss
```
Esperado: `Successful compile` y `Output/KNOB_App_Installer.exe`.

- [ ] **Paso 5: Probar el build**

Con la v3.3 cerrada (bandeja → Salir), ejecutar `dist/KNOB/KNOB.exe`. Verificar que abre en menos de 2 s, se ve igual que en desarrollo (logo con la fuente Amazing Views) y conecta con el KNOB. Revisar `dist/KNOB/knob_debug.log` si algo falla.

- [ ] **Paso 6: Commit**

```bash
git add KNOB.spec KNOB_App_Installer.iss release.py version.txt
git commit -m "build: empaquetado v4.0.0 en carpeta con PySide6 e instalador actualizado" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Tarea 13: Validación con el hardware y limpieza

**Archivos:**
- Borrar (solo con el OK del usuario): `app.py`
- Modificar: `requirements.txt` (no cambia: ya no lista customtkinter ni pystray)

- [ ] **Paso 1: Lista de pruebas manuales con el usuario**

Instalar `Output/KNOB_App_Installer.exe` encima de la v3.3 instalada y recorrer con el usuario la lista de la sección 9 del spec:
1. Con el KNOB conectado, "Buscando" pasa a "Conectado" en menos de 3 s.
2. Cada fader sigue suave y cambia el volumen de la app correcta (comparar con el mezclador de Windows).
3. Los VU se mueven con el audio y escalan con el fader.
4. Encoders: giro, click, doble click, mantener y girar (incluido el cambio de salida de audio).
5. Grabar un atajo de teclado y ejecutarlo.
6. Desenchufar y reenchufar: reconecta solo.
7. Con el Serial Monitor de Arduino abierto aparece la pastilla roja.
8. Cambiar perfil y tema; crear, renombrar y borrar perfil; reiniciar y verificar que persiste.
9. "App en foco" muestra la app activa y respeta "el último que se movió gana".
10. Bandeja, `--minimized`, iniciar con Windows y segunda instancia bloqueada.
11. Los perfiles y el tema de la v3.3 se conservan después de instalar.
12. Escala de Windows al 100%, 125% y 150%: nada se corta.

Anotar cada falla, arreglarla con su test cuando sea lógica del core y repetir la prueba afectada.

- [ ] **Paso 2: Borrar la app vieja (con OK explícito del usuario)**

```bash
git rm app.py
venv/Scripts/python.exe -m pytest
git commit -m "chore: quitar app.py (customtkinter); la v4 la reemplaza" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

- [ ] **Paso 3: Integrar**

Usar superpowers:finishing-a-development-branch para decidir con el usuario cómo llevar `rediseno-qt` a `master`. El release (`python release.py v4.0.0 TOKEN`) lo ejecuta el usuario, porque necesita su token de GitHub.

---

## Nota para el usuario: "mantener y girar" en v3.3

Al portar la lógica apareció un comportamiento raro de v3.3. El plan lo **conserva tal cual** (Tarea 3, `test_mantener_y_girar_comportamiento_v33`). Corregirlo queda como decisión aparte:

- Si un encoder tiene "Mantener y girar" en Volumen general, Micrófono o App en foco, mantener el botón **no** controla ese destino con el giro. Le aplica el valor congelado de la perilla. Por ejemplo, con tu config (los dos encoders con `click_turn: mic`), mantener el botón izquierdo pone el micrófono al nivel del volumen general.
- Al soltar, la perilla salta a la posición girada, porque el firmware sigue contando.
- "Cambiar salida de audio" sí funciona bien, porque usa el valor real del hardware.

Para arreglarlo bien conviene un cambio chico en el firmware: que no cuente el giro mientras el botón está mantenido, o que mande el delta aparte. Hacerlo solo en la app no alcanza.
