import os
import sys
import json
import winreg
import subprocess
import tempfile
import urllib.request
import customtkinter
import tkinter as tk
import serial
import threading
import time
import serial.tools.list_ports
from PIL import Image, ImageDraw
import math
import ctypes
import psutil
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume, IAudioMeterInformation
from comtypes import CLSCTX_ALL
import comtypes

# --- SINGLE INSTANCE LOCK ---
_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "Global\\KNOBAppMutex")
if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
    ctypes.windll.user32.MessageBoxW(0, "KNOB ya está en ejecución.", "KNOB", 0x30)
    sys.exit(0)

# --- DEBUG LOG (writes to knob_debug.log next to the exe/script) ---
_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, "frozen", False) else __file__)), "knob_debug.log")
def _log(msg):
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as _f:
            _f.write(f"[{time.strftime('%H:%M:%S')}] {msg}\n")
    except Exception:
        pass

# --- DPI: make the app DPI-unaware so the OS handles scaling as a bitmap.
# This keeps the absolute place() layout intact at all Windows scale settings.
# Must be called before any window is created.
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)   # DPI unaware — OS scales
except Exception:
    pass

try:
    import pystray
    PYSTRAY_AVAILABLE = True
except ImportError:
    PYSTRAY_AVAILABLE = False

# --- APP ICON ---
_ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
_app_icon_pil = Image.open(_ICON_PATH).convert("RGBA")

# --- CUSTOM DIAL WIDGET (270° potentiometer-style arc) ---
class DialWidget(tk.Canvas):
    """Potentiometer-style dial: 0 at 7-o'clock, max at 5-o'clock, 270° CW sweep."""
    _START_DEG = 225   # angle (CCW from East, y-up) of minimum position (7 o'clock)
    _SWEEP_DEG = 270   # total clockwise sweep in degrees

    def __init__(self, master, start=0, end=100, radius=50, unit_length=15,
                 unit_width=3, needle_color="white",
                 color_gradient=("#B0B0B0", "#363636"), text="", bg="black", **kw):
        self._r = radius
        self._ul = unit_length
        self._uw = unit_width
        self._needle_color = needle_color
        self._grad = color_gradient
        self._start_val = start
        self._end_val = end
        size = (radius + unit_length) * 2 + 4
        self._cx = size // 2
        self._cy = size // 2
        super().__init__(master, width=size, height=size, bg=bg,
                         bd=0, highlightthickness=0, **kw)
        self._needle_id = None
        self.value = start
        self._draw_arc()
        self._draw_ticks()
        self._draw_needle()

    def _lerp_color(self, c1, c2, t):
        r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
        r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
        return "#{:02x}{:02x}{:02x}".format(
            int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t), int(b1 + (b2 - b1) * t))

    def _angle_deg(self, val):
        frac = (val - self._start_val) / max(self._end_val - self._start_val, 1)
        return self._START_DEG - frac * self._SWEEP_DEG  # CW = decreasing angle

    def _xy(self, r, angle_deg):
        rad = math.radians(angle_deg)
        return self._cx + r * math.cos(rad), self._cy - r * math.sin(rad)

    def _draw_arc(self):
        cx, cy, r = self._cx, self._cy, self._r
        # start=-45 (5 o'clock), extent=270 CCW → through 12, 9 o'clock, to 7 o'clock
        self.create_arc(cx - r, cy - r, cx + r, cy + r,
                        start=-45, extent=270, style="arc",
                        outline="#444444", width=2)

    def _draw_ticks(self):
        n = 37  # 36 intervals over 270°
        for i in range(n):
            angle = self._START_DEG - (i / (n - 1)) * self._SWEEP_DEG
            x1, y1 = self._xy(self._r, angle)
            x2, y2 = self._xy(self._r + self._ul, angle)
            color = self._lerp_color(self._grad[0], self._grad[1], i / (n - 1))
            self.create_line(x1, y1, x2, y2, fill=color, width=self._uw)

    def _draw_needle(self):
        angle = self._angle_deg(self.value)
        x, y = self._xy(self._r, angle)
        if self._needle_id is None:
            self._needle_id = self.create_line(
                self._cx, self._cy, x, y,
                fill=self._needle_color, width=self._uw + 1)
        else:
            self.coords(self._needle_id, self._cx, self._cy, x, y)

    def set(self, value):
        self.value = max(self._start_val, min(self._end_val, value))
        self._draw_needle()

    def get(self):
        return self.value

    def configure(self, **kw):
        if "bg" in kw:
            super().configure(bg=kw.pop("bg"))
        if kw:
            super().configure(**kw)

# --- VU METER WIDGET ---
class VUMeter(tk.Canvas):
    """Segmented LED-style vertical VU meter."""
    SEG_H   = 6    # height of each colored segment
    SEG_GAP = 2    # gap between segments (background color)

    def __init__(self, master, width=8, height=400, bg="black", **kw):
        super().__init__(master, width=width, height=height, bg=bg,
                         bd=0, highlightthickness=0, **kw)
        self._h = height
        self._width = width
        self._bar_id  = self.create_rectangle(0, height, width, height,
                                              fill="#00C853", outline="")
        self._peak_id = self.create_line(0, height, width, height,
                                         fill="white", width=2)
        # Segment separators drawn on top — create_line after bar so z-order is correct
        self._segs = []
        for y in range(0, height, self.SEG_H + self.SEG_GAP):
            lid = self.create_line(0, y, width, y, fill=bg, width=self.SEG_GAP)
            self._segs.append(lid)

    def set_level(self, level: float, peak: float):
        h, w = self._h, self._width
        bar_h  = max(0, int(level * h))
        y_top  = h - bar_h
        color  = ("#F44336" if level > 0.85 else
                  "#FFD600" if level > 0.70 else "#00C853")
        if bar_h > 0:
            self.coords(self._bar_id, 0, y_top, w, h)
            self.itemconfig(self._bar_id, fill=color)
        else:
            self.coords(self._bar_id, 0, h, w, h)
        peak_y = max(0, int((1.0 - peak) * h) - 1)
        if peak > 0.02:
            self.coords(self._peak_id, 0, peak_y, w, peak_y)
        else:
            self.coords(self._peak_id, 0, h, w, h)

    def update_bg(self, color: str):
        self.configure(bg=color)
        for lid in self._segs:
            self.itemconfig(lid, fill=color)

# --- CONFIG ---
# Use %APPDATA%\KNOB so config is always writable even when installed
# in a protected directory like C:\Program Files.
_APPDATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "KNOB")
os.makedirs(_APPDATA_DIR, exist_ok=True)
CONFIG_PATH = os.path.join(_APPDATA_DIR, "config.json")
_log(f"CONFIG_PATH={CONFIG_PATH}  exists={os.path.exists(CONFIG_PATH)}")

def load_config():
    try:
        with open(CONFIG_PATH) as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(data):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

config = load_config()
# Migrate flat assignments to profile structure on first run
if "profiles" not in config:
    config["profiles"] = {
        "Default": {
            "channel_assignments": config.get("channel_assignments", {}),
            "encoder_assignments": config.get("encoder_assignments", [{}, {}])
        }
    }
    config["active_profile"] = "Default"
    save_config(config)

# --- THEMES ---
# Each palette: bg (dark bg), bg_mid (panel/frame bg), accent, accent_hover,
#               btn_un (unassigned/inactive button), text, text_dim
THEMES = {
    "Magma":   {"accent": "#B31D1D", "accent_hover": "#8A1010", "bg": "#0D0404",
               "bg_mid": "#1E0808", "text": "#FFFFFF", "text_dim": "#FFAAAA", "btn_un": "#640606"},
    "Cobalt":   {"accent": "#2533B0", "accent_hover": "#1A228A", "bg": "#04060D",
               "bg_mid": "#1B368F", "text": "#FFFFFF", "text_dim": "#AABBFF", "btn_un": "#0F1E74"},
    "Onyx":  {"accent": "#555555", "accent_hover": "#6A6A6A", "bg": "#000000",
               "bg_mid": "#111111", "text": "#DDDDDD", "text_dim": "#999999", "btn_un": "#1A1A1A"},
    "Arctic": {"accent": "#888888", "accent_hover": "#606060", "bg": "#F0F0F0",
               "bg_mid": "#DEDEDE", "text": "#111111", "text_dim": "#555555", "btn_un": "#BBBBBB"},
    "Steel":   {"accent": "#A6A6A6", "accent_hover": "#878787", "bg": "#0D0D0D",
               "bg_mid": "#858585", "text": "#FFFFFF", "text_dim": "#CCCCCC", "btn_un": "#616060"},
    "Nebula": {"accent": "#621694", "accent_hover": "#4A1070", "bg": "#08020F",
               "bg_mid": "#4E17A1", "text": "#FFFFFF", "text_dim": "#DDAAFF", "btn_un": "#350569"},
}

_default_theme = next(iter(THEMES))
current_theme = config.get("theme", _default_theme)
if current_theme not in THEMES:
    current_theme = _default_theme
active_tab = "knob"

# --- ROOT ---
customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("dark-blue")
# Disable CTk's auto-DPI scaling so it doesn't fight with the OS bitmap scaling
customtkinter.set_widget_scaling(1.0)
customtkinter.set_window_scaling(1.0)

root = customtkinter.CTk()
root.title("KNOB App")
root.geometry("700x800")
root.resizable(False, False)

# Taskbar / title-bar icon — convert PNG→ICO in temp dir for reliable Windows support
import tempfile as _tempfile
try:
    _ico_path = os.path.join(_tempfile.gettempdir(), "knob_app.ico")
    _app_icon_pil.resize((256, 256), Image.LANCZOS).save(
        _ico_path, format="ICO",
        sizes=[(256, 256), (64, 64), (32, 32), (16, 16)]
    )
    root.iconbitmap(_ico_path)
except Exception:
    pass

# Register AmazingViews.ttf with Windows GDI so Tk can use it by family name.
# Works both from source (fonts/ next to app.py) and from PyInstaller exe (_MEIPASS).
def _register_font(filename):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(base, "fonts", filename)
    if os.path.exists(path):
        ctypes.windll.gdi32.AddFontResourceW(path)

_register_font("AmazingViews.ttf")
mi_fuente_personalizada = customtkinter.CTkFont(family="Amazing Views", size=70)

root.grid_columnconfigure(0, weight=0)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=1)

# --- LEFT PANEL (custom vertical tabs) ---
left_panel = customtkinter.CTkFrame(root, width=140, corner_radius=8)
left_panel.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
left_panel.grid_propagate(False)

tab_btn_knob = customtkinter.CTkButton(
    left_panel, text="KNOB", height=40, corner_radius=6,
    command=lambda: switch_tab("knob")
)
tab_btn_knob.pack(fill="x", padx=8, pady=(12, 3))

tab_btn_config = customtkinter.CTkButton(
    left_panel, text="Config", height=40, corner_radius=6,
    command=lambda: switch_tab("config")
)
tab_btn_config.pack(fill="x", padx=8, pady=(3, 10))

content_knob = customtkinter.CTkFrame(left_panel, fg_color="transparent")
content_config = customtkinter.CTkFrame(left_panel, fg_color="transparent")
content_knob.pack(fill="both", expand=True, padx=4, pady=4)  # KNOB shown by default

def _read_version():
    try:
        if getattr(sys, "frozen", False):
            # Primero busca en el directorio del exe (donde lo copia el instalador)
            exe_dir = os.path.dirname(sys.executable)
            path = os.path.join(exe_dir, "version.txt")
            if not os.path.exists(path):
                path = os.path.join(sys._MEIPASS, "version.txt")
        else:
            path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.txt")
        with open(path, encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return "v?.?.?"
APP_VERSION = _read_version()
_RELEASES_URL = "https://api.github.com/repos/fpesse/KNOB-app/releases/latest"

btn_version = customtkinter.CTkLabel(
    left_panel, text=APP_VERSION,
    font=customtkinter.CTkFont(size=10),
    cursor="arrow"
)
btn_version.pack(side="bottom", pady=(0, 8))

def _parse_version(v):
    try:
        return tuple(int(x) for x in v.lstrip("v").split("."))
    except Exception:
        return (0,)

def _check_for_update():
    try:
        req = urllib.request.Request(
            _RELEASES_URL,
            headers={"User-Agent": "KNOB-App", "Accept": "application/vnd.github+json"}
        )
        with urllib.request.urlopen(req, timeout=8) as r:
            data = json.loads(r.read())
        latest_tag  = data.get("tag_name", "")
        asset_url   = next(
            (a["browser_download_url"] for a in data.get("assets", [])
             if a["name"].endswith(".exe")), None
        )
        if not asset_url:
            return
        if _parse_version(latest_tag) > _parse_version(APP_VERSION):
            root.after(0, lambda: _show_update_available(latest_tag, asset_url))
    except Exception:
        pass

def _show_update_available(latest_tag, asset_url):
    t = THEMES[current_theme]
    btn_version.configure(
        text=f"↑ {latest_tag} disponible",
        fg_color=t["accent"],
        text_color=t["text"], cursor="hand2"
    )
    btn_version.bind("<Button-1>", lambda e: _start_download(latest_tag, asset_url))

def _start_download(latest_tag, asset_url):
    t = THEMES[current_theme]
    btn_version.unbind("<Button-1>")
    btn_version.configure(text="Descargando 0%",
                          fg_color=t["btn_un"], cursor="arrow")
    def download():
        try:
            tmp = tempfile.mktemp(suffix=".exe", prefix="KNOB_update_")
            with urllib.request.urlopen(asset_url, timeout=60) as r:
                total = int(r.headers.get("Content-Length", 0))
                downloaded = 0
                chunk = 65536
                with open(tmp, "wb") as f:
                    while True:
                        buf = r.read(chunk)
                        if not buf:
                            break
                        f.write(buf)
                        downloaded += len(buf)
                        if total:
                            pct = int(downloaded * 100 / total)
                            root.after(0, lambda p=pct: btn_version.configure(
                                text=f"Descargando {p}%"))
            root.after(0, lambda: _launch_installer(tmp))
        except Exception as e:
            root.after(0, lambda: btn_version.configure(text="Error al descargar"))
    threading.Thread(target=download, daemon=True).start()

def _launch_installer(path):
    os.startfile(path)
    root.after(500, on_close)

# KNOB tab: connection status label
lbl_status_knob = customtkinter.CTkLabel(
    content_knob, text="Buscando\nKNOB...",
    font=customtkinter.CTkFont(size=11), justify="center"
)
lbl_status_knob.pack(pady=(20, 0))

lbl_audio_device = customtkinter.CTkLabel(
    content_knob, text="",
    font=customtkinter.CTkFont(size=9), justify="center",
    wraplength=120
)
lbl_audio_device.pack(pady=(4, 0))

def _update_tab_buttons():
    t = THEMES[current_theme]
    tab_btn_knob.configure(
        fg_color=t["accent"] if active_tab == "knob" else t["btn_un"],
        hover_color=t["accent_hover"],
        text_color=t["text"]
    )
    tab_btn_config.configure(
        fg_color=t["accent"] if active_tab == "config" else t["btn_un"],
        hover_color=t["accent_hover"],
        text_color=t["text"]
    )

def switch_tab(tab):
    global active_tab
    active_tab = tab
    if tab == "knob":
        content_config.pack_forget()
        content_knob.pack(fill="both", expand=True, padx=4, pady=4)
    else:
        content_knob.pack_forget()
        content_config.pack(fill="both", expand=True, padx=4, pady=4)
    _update_tab_buttons()

# --- RIGHT FRAME ---
frame_derecho = customtkinter.CTkFrame(root, corner_radius=8)
frame_derecho.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

icono_estado = customtkinter.CTkLabel(frame_derecho, text="", width=40, height=40)
icono_estado.place(x=465, y=10)

label2 = customtkinter.CTkLabel(frame_derecho, text="KNOB", font=mi_fuente_personalizada)
label2.place(x=160, y=20)

# --- ICON FUNCTIONS ---
def crear_icono_busqueda(angle=0):
    img = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, 36, 36), fill="#FFA500", outline="#FFA500")
    cx, cy, r = 20, 20, 13
    x1 = cx + r * math.cos(math.radians(angle))
    y1 = cy + r * math.sin(math.radians(angle))
    x2 = cx + (r - 7) * math.cos(math.radians(angle + 20))
    y2 = cy + (r - 7) * math.sin(math.radians(angle + 20))
    draw.line((cx, cy, x1, y1), fill="white", width=4)
    draw.polygon([(x1, y1), (x2, y2), (cx, cy)], fill="white")
    return customtkinter.CTkImage(light_image=img, dark_image=img, size=(40, 40))

def crear_icono_check():
    img = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, 36, 36), fill="#00C853", outline="#00C853")
    draw.line((12, 22, 19, 29, 29, 13), fill="white", width=5, joint="curve")
    return customtkinter.CTkImage(light_image=img, dark_image=img, size=(40, 40))

# --- SERIAL / CONNECTION ---
conectado = False
icono_angle = 0
ser_global = None
_conn_gen = 0   # increments on each successful connect; leer_datos exits if gen changes
latest_vals = None
vals_lock = threading.Lock()

def animar_busqueda():
    global icono_angle
    if not conectado:
        img = crear_icono_busqueda(icono_angle)
        icono_estado.configure(image=img)
        icono_estado.image = img
        icono_angle = (icono_angle + 30) % 360
        icono_estado.after(100, animar_busqueda)

def _ui_set_conectado():
    img = crear_icono_check()
    icono_estado.configure(image=img)
    icono_estado.image = img
    lbl_status_knob.configure(text="Conectado", text_color="#00C853")

def _ui_set_buscando():
    icono_estado.configure(text="", image=None)
    lbl_status_knob.configure(text="Buscando KNOB...",
                               text_color=THEMES[current_theme]["text_dim"])

# --- SMOOTHING ---
# smooth_vals: current interpolated values (0-1023 range)
# target_vals: latest raw values from Arduino
SMOOTH_ALPHA    = 1   # EMA speed: 0.5 → halves distance every 16ms (~48ms to settle)
DEADBAND        = 10     # ADC noise tolerance — ignore changes smaller than ±3 units
MUTE_THRESHOLD  = 15    # raw values ≤ this → apply 0 (true mute, avoids lingering 1-2)

def raw_to_level(raw_val: float) -> float:
    """Convert raw ADC value (0-1023) to audio level using a logarithmic curve.
    x^2 matches human loudness perception: fine control at low volumes, natural feel at high."""
    if raw_val <= MUTE_THRESHOLD:
        return 0.0
    linear = max(0.0, min(1.0, raw_val / 1023.0))
    return linear * linear   # square law ≈ perceptual loudness
smooth_vals = [512.0] * 5
target_vals = [512.0] * 5
_dial_muted    = [False, False]   # True when dial is muted via button press
_dial_saved_val = [512.0, 512.0]  # smooth_val saved before button-mute

_last_vol_ms = 0   # timestamp of last volume API call (monotonic ms)
VOL_RATE_MS   = 40  # call audio API at most every 40ms (25Hz) to avoid blocking UI

# ── Focus priority tracking ("last mover wins") ────────────────────────────
# Maps app_name → "direct" | "focus". Whichever moved last controls that app.
_channel_priority: dict = {}
_last_raw_per_channel: list = [None] * 5   # last raw_val seen per channel
_focus_took_at: dict = {}   # app → raw_val of its direct channel when focus took priority
_focus_prev_raw: float = None  # last raw_val of the focus channel (to detect actual movement)

# ── Audio output device switching ─────────────────────────────────────────
_output_devices   = []        # [(device_id, friendly_name), ...]
_output_dev_idx   = 0         # index of current default device
_dev_switch_ms    = 0         # debounce timestamp
_enc_prev_raw     = [None, None]  # previous encoder value per encoder (for direction)
_enc_hold_saved   = [None, None]  # smooth_val saved when button is pressed (for restore on release)

# ── Focus tracking ─────────────────────────────────────────────────────────
_last_focus_exe: str = ""     # last non-KNOB foreground exe

def _focus_tracker():
    """Background thread: polls foreground window every 150ms, caches last non-KNOB exe."""
    global _last_focus_exe
    own_pid = os.getpid()
    while True:
        try:
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if hwnd:
                pid = ctypes.c_ulong()
                ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value != own_pid:
                    exe = strip_exe(psutil.Process(pid.value).name()).capitalize()
                    if exe:
                        _last_focus_exe = exe
        except Exception:
            pass
        time.sleep(0.15)

threading.Thread(target=_focus_tracker, daemon=True).start()

def _get_device_friendly_name(device_id: str) -> str:
    """Build full name like 'Speakers (UG-08S)' from registry properties."""
    try:
        guid = device_id.rsplit('.', 1)[-1]
        key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{guid}\Properties"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0,
                             winreg.KEY_READ | 0x0100)
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

def _load_output_devices():
    global _output_devices, _output_dev_idx
    try:
        enumerator = AudioUtilities.GetDeviceEnumerator()
        collection = enumerator.EnumAudioEndpoints(0, 1)  # eRender=0, ACTIVE=1
        devs = []
        for i in range(collection.GetCount()):
            dev = collection.Item(i)
            dev_id = dev.GetId()
            name = _get_device_friendly_name(dev_id) or f"Dispositivo {i+1}"
            devs.append((dev_id, name))
        _output_devices = devs
        try:
            default_id = enumerator.GetDefaultAudioEndpoint(0, 1).GetId()
            for i, (did, name) in enumerate(_output_devices):
                if did == default_id:
                    _output_dev_idx = i
                    root.after(0, lambda n=name: _update_device_label(n))
                    break
        except Exception:
            pass
    except Exception as e:
        _log(f"_load_output_devices: {e}")

def _set_default_output_device(device_id: str):
    """Switch default audio output via undocumented IPolicyConfig COM interface."""
    try:
        from comtypes import GUID as _GUID
        clsid      = _GUID('{870AF99C-171D-4F9E-AF0D-E63DF40C2BC9}')
        iid_unk    = _GUID('{00000000-0000-0000-C000-000000000046}')
        iid_policy = _GUID('{F8679F50-850A-41CF-9C72-430F290290C8}')
        # 1. Create COM object
        unk_ptr = ctypes.POINTER(ctypes.c_void_p)()
        hr = ctypes.windll.ole32.CoCreateInstance(
            ctypes.byref(clsid), None, 1,
            ctypes.byref(iid_unk), ctypes.byref(unk_ptr))
        if hr != 0:
            return
        # 2. QI for IPolicyConfig to get the correct vtable
        vtable_unk = ctypes.cast(unk_ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
        QI = ctypes.WINFUNCTYPE(
            ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p))(vtable_unk[0])
        policy_ptr = ctypes.c_void_p()
        hr = QI(ctypes.cast(unk_ptr, ctypes.c_void_p).value,
                ctypes.addressof(iid_policy), ctypes.byref(policy_ptr))
        if hr != 0 or not policy_ptr:
            return
        # 3. Call SetDefaultEndpoint (vtable index 13)
        vtable = ctypes.cast(policy_ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
        SetDefaultEndpoint = ctypes.WINFUNCTYPE(
            ctypes.HRESULT, ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint)(vtable[13])
        for role in range(3):
            SetDefaultEndpoint(policy_ptr.value, device_id, role)
        ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtable[2])(policy_ptr.value)
    except Exception as e:
        _log(f"_set_default_output_device: {e}")

def _update_device_label(name: str, highlight: bool = False):
    lbl_audio_device.configure(text=name)
    if highlight:
        lbl_audio_device.configure(text_color="#FFFFFF")
        root.after(1500, lambda: lbl_audio_device.configure(
            text_color=THEMES[current_theme]["text_dim"]))

def _cycle_output_device(direction: int):
    global _output_dev_idx, _dev_switch_ms
    if not _output_devices:
        return
    now_ms = time.monotonic() * 1000
    if now_ms - _dev_switch_ms < 800:
        return
    _dev_switch_ms = now_ms
    _output_dev_idx = (_output_dev_idx + direction) % len(_output_devices)
    dev_id, dev_name = _output_devices[_output_dev_idx]
    threading.Thread(target=_set_default_output_device, args=(dev_id,), daemon=True).start()
    root.after(0, lambda n=dev_name: _update_device_label(n, highlight=True))

def _show_device_toast(name: str):
    pass  # replaced by sidebar label highlight

def _apply_volume_by_key(key, raw_val):
    level = raw_to_level(raw_val)
    try:
        if key == "master":
            set_master_volume(level)
        elif key == "mic":
            set_mic_volume(level)
        elif key == "focus":
            set_focus_app_volume(level)
        else:
            set_app_volume(key, level)
    except Exception as e:
        print(f"[DEBUG] click_turn volume: {e}")

def _apply_ui_from_smooth():
    global _last_vol_ms
    # UI widgets update every frame (60 fps) — pure canvas/widget ops, very fast
    slider1.set(smooth_vals[0])
    slider2.set(smooth_vals[1])
    slider3.set(smooth_vals[2])
    # Freeze dial UI while held with a non-none click_turn action
    _knob_val_0 = _enc_hold_saved[0] if (encoder_held[0] and encoder_assignments[0].get("click_turn","none") != "none" and _enc_hold_saved[0] is not None) else smooth_vals[3]
    _knob_val_1 = _enc_hold_saved[1] if (encoder_held[1] and encoder_assignments[1].get("click_turn","none") != "none" and _enc_hold_saved[1] is not None) else smooth_vals[4]
    knob_izquierda.set(_knob_val_0 / 1023 * 100)
    knob_derecha.set(_knob_val_1 / 1023 * 100)
    # Volume API (pycaw COM calls) throttled to 25 Hz to keep main thread free
    now_ms = time.monotonic() * 1000
    if now_ms - _last_vol_ms < VOL_RATE_MS:
        return
    _last_vol_ms = now_ms
    for i in range(5):
        # Dials (i=3 → enc 0, i=4 → enc 1): if held, redirect to click_turn channel
        if i >= 3:
            enc_idx = i - 3
            if encoder_held[enc_idx]:
                turn_key = encoder_assignments[enc_idx].get("click_turn", "none")
                if turn_key == "output_device":
                    curr = target_vals[i]   # raw hardware value, not frozen
                    prev = _enc_prev_raw[enc_idx]
                    if prev is not None and abs(curr - prev) > 35:
                        _cycle_output_device(1 if curr > prev else -1)
                    _enc_prev_raw[enc_idx] = curr
                    continue
                elif turn_key == "none":
                    _enc_prev_raw[enc_idx] = None
                else:
                    _enc_prev_raw[enc_idx] = None
                    _apply_volume_by_key(turn_key, smooth_vals[i])
                    continue
        apply_volume(i, smooth_vals[i])
    # Refresh mute state on buttons if any channel crossed the mute threshold
    for i in range(5):
        now = bool(channel_assignments.get(i)) and smooth_vals[i] <= MUTE_THRESHOLD
        if now != _prev_muted[i]:
            update_channel_buttons()
            break

def leer_datos(gen):
    global conectado, ser_global, latest_vals
    ser_global.timeout = 0.05
    prev_btn = [0, 0]   # last received button value per encoder (edge detection)
    while conectado and ser_global and ser_global.is_open and _conn_gen == gen:
        try:
            linea = ser_global.readline().decode(errors="ignore").strip()
            if not linea:
                continue
            partes = linea.split("|")
            if len(partes) >= 5:
                vals = [int(x) for x in partes[:5]]
                with vals_lock:
                    latest_vals = vals
                # Arduino sends: v0|v1|v2|v3|v4|btnL|btnR
                # btnX: 0=none, 1=click, 2=dblclick, 3=held_start, 4=held_end
                # Edge detection: only fire when value changes from previous frame
                for enc_idx, col in enumerate([5, 6]):
                    if len(partes) > col:
                        try:
                            ev = int(partes[col])
                            if ev != prev_btn[enc_idx]:
                                prev_btn[enc_idx] = ev
                                if ev > 0:
                                    root.after(0, lambda e=enc_idx, c=ev: handle_button_event(e, c))
                        except ValueError:
                            pass
        except Exception as e:
            print(f"[DEBUG] Error leyendo datos: {e}")
            break

    # Only the active generation triggers a reconnect — prevents duplicate search threads
    if _conn_gen == gen and conectado:
        print("[DEBUG] Conexion perdida, reiniciando busqueda...")
        conectado = False
        try:
            ser_global.close()
        except Exception:
            pass
        ser_global = None
        root.after(0, _ui_set_buscando)
        root.after(0, animar_busqueda)
        threading.Thread(target=buscar_y_handshake, daemon=True).start()

def poll_ui():
    global latest_vals
    with vals_lock:
        vals = latest_vals
        latest_vals = None
    if vals is not None:
        for i in range(5):
            # If this dial is button-muted, ignore physical movement
            if i >= 3 and _dial_muted[i - 3]:
                target_vals[i] = 0.0
                continue
            new = float(vals[i])
            # Deadband: only update target when the physical change is real (not ADC noise)
            if abs(new - target_vals[i]) >= DEADBAND:
                target_vals[i] = new
    # EMA: move smooth_vals toward target_vals each frame
    needs_update = False
    for i in range(5):
        # Keep muted dials pinned at 0
        if i >= 3 and _dial_muted[i - 3]:
            if smooth_vals[i] != 0.0:
                smooth_vals[i] = 0.0
                needs_update = True
            continue
        # Freeze encoder dial while held with a non-none click_turn action
        if i >= 3:
            enc_idx = i - 3
            try:
                held = encoder_held[enc_idx]
                ct_key = encoder_assignments[enc_idx].get("click_turn", "none")
            except Exception:
                held = False
                ct_key = "none"
            if held and ct_key != "none":
                saved = _enc_hold_saved[enc_idx]
                if saved is not None:
                    smooth_vals[i] = saved  # freeze UI/audio only, NOT target_vals
                needs_update = True  # force _apply_ui_from_smooth while held
                continue
        diff = target_vals[i] - smooth_vals[i]
        if abs(diff) < 0.5:
            if smooth_vals[i] != target_vals[i]:
                smooth_vals[i] = target_vals[i]   # snap when basically there
                needs_update = True
        else:
            smooth_vals[i] += diff * SMOOTH_ALPHA
            needs_update = True
    if needs_update:
        _apply_ui_from_smooth()
    root.after(16, poll_ui)

def _try_handshake(device_name):
    """Open device_name and complete handshake. Returns (Serial, busy_flag) tuple."""
    try:
        ser = serial.Serial(device_name, 115200, timeout=0.5, write_timeout=1.0)
    except Exception as e:
        busy = "Access is denied" in str(e) or "PermissionError" in str(e)
        _log(f"{device_name}: no se pudo abrir ({'OCUPADO' if busy else str(e)})")
        return None, busy

    _log(f"{device_name}: puerto abierto, esperando respuesta Arduino...")
    t_end      = time.time() + 12.0
    # last_probe starts at now so we read first before sending anything.
    # If Arduino is already streaming (app restart) readline() catches it immediately
    # without ever blocking on write(). Probe is sent only after 1s with no data.
    last_probe = time.time()
    try:
        while time.time() < t_end:
            # Read first — catches streaming data or AWAIT_INIT immediately
            line = ser.readline().decode(errors="ignore").strip()
            if line:
                _log(f"{device_name}: recibido: '{line}'")
                if line == "{KNOB_CMD_AWAIT_INIT}":
                    ser.write(b"{KNOB_CMD_START}\n")
                    last_probe = time.time()
                elif line == "{KNOB_LOG_INITIALIZED}":
                    _log(f"{device_name}: CONECTADO (handshake OK)")
                    return ser, False
                elif line.count("|") >= 4:
                    _log(f"{device_name}: CONECTADO (ya en streaming)")
                    return ser, False
            else:
                # No data yet — send probe every 1s to wake up AWAIT state
                now = time.time()
                if now - last_probe >= 1.0:
                    try:
                        ser.write(b"{KNOB_CMD_START}\n")
                        _log(f"{device_name}: probe enviado")
                    except serial.SerialTimeoutException:
                        _log(f"{device_name}: write timeout")
                    last_probe = now
    except Exception as e:
        _log(f"{device_name}: excepcion: {e}")

    _log(f"{device_name}: timeout/fallo, cerrando")
    try:
        ser.close()
    except Exception:
        pass
    return None, False


def _connect_on_port(port_device):
    """Try handshake on port_device; on success set globals and return (True, False).
    Returns (False, True) if port was busy."""
    global conectado, ser_global, _conn_gen
    ser, busy = _try_handshake(port_device)
    if busy:
        return False, True
    if ser:
        _conn_gen += 1
        gen = _conn_gen
        conectado = True
        ser_global = ser
        cfg = load_config()
        cfg["last_port"] = port_device
        save_config(cfg)
        _log(f"Conectado en {port_device} (gen={gen})")
        root.after(0, _ui_set_conectado)
        threading.Thread(target=leer_datos, args=(gen,), daemon=True).start()
        return True, False
    return False, False

def buscar_y_handshake():
    global conectado, ser_global
    try:
        while not conectado:
            last_port = load_config().get("last_port")
            todos = list(serial.tools.list_ports.comports())
            _log(f"Puertos: {[(p.device, p.description, p.vid) for p in todos]}")

            # Only try real Arduino/USB serial ports — never bare COM ports like COM1
            candidatos = [p for p in todos
                          if p.vid is not None
                          or any(kw in (p.description or "").upper()
                                 for kw in ("USB", "CH340", "CP210", "FTDI", "ARDUINO"))]
            candidatos = sorted(candidatos, key=lambda p: (0 if p.device == last_port else 1))

            if not candidatos:
                _log("No hay puertos Arduino/USB disponibles")
                root.after(0, _ui_set_buscando)
                time.sleep(2)
                continue

            found = False
            any_busy = False
            for p in candidatos:
                _log(f"Probando: {p.device} ({p.description})")
                ok, busy = _connect_on_port(p.device)
                if ok:
                    found = True
                    break
                if busy:
                    any_busy = True
                    root.after(0, lambda d=p.device: lbl_status_knob.configure(
                        text=f"Cerrar Serial Monitor\n({d} ocupado)",
                        text_color=THEMES[current_theme]["text_dim"]))

            if not found:
                _log("KNOB no encontrado, reintentando en 2s...")
                if not any_busy:
                    root.after(0, _ui_set_buscando)
                time.sleep(2)
    except Exception as e:
        _log(f"ERROR en buscar_y_handshake: {e}")

def iniciar_conexion():
    threading.Thread(target=buscar_y_handshake, daemon=True).start()
    animar_busqueda()
    poll_ui()

iniciar_conexion()

# --- DIALS ---
# Sliders are on root at x=340,450,560 (center=352,462,572).
# frame_derecho inner x offset from root ≈ 172px.
# Slider centers in frame coords: ~180, ~290, ~400.
# Dial widget (radius=50, unit_length=15): circle center at ~x+65.
# Aligning dial centers with sliders 1 and 3: x=115 (center≈180), x=335 (center≈400).
knob_izquierda = DialWidget(master=frame_derecho,
                            radius=50, unit_length=15, unit_width=3,
                            color_gradient=("#B0B0B0", "#363636"),
                            needle_color="white")
knob_izquierda.place(x=65, y=615)

knob_derecha = DialWidget(master=frame_derecho,
                          radius=50, unit_length=15, unit_width=3,
                          color_gradient=("#B0B0B0", "#363636"),
                          needle_color="white")
knob_derecha.place(x=285, y=615)

# --- SLIDERS ---
slider1 = customtkinter.CTkSlider(root,
    from_=0, to=1023,
    orientation="vertical", width=25, height=400,
    bg_color="black", button_color="gray20")
slider1.place(x=308, y=160)

slider2 = customtkinter.CTkSlider(root,
    from_=0, to=1023,
    orientation="vertical", width=25, height=400,
    bg_color="black", button_color="gray20")
slider2.place(x=418, y=160)

slider3 = customtkinter.CTkSlider(root,
    from_=0, to=1023,
    orientation="vertical", width=25, height=400,
    bg_color="black", button_color="gray20")
slider3.place(x=528, y=160)

# Block all mouse interaction on sliders — hardware-only control
def _block_mouse(_event): return "break"
_BLOCK_EVENTS = ("<Button-1>", "<B1-Motion>", "<ButtonRelease-1>",
                 "<Button-2>", "<B2-Motion>", "<Button-3>", "<MouseWheel>")
for _sl in (slider1, slider2, slider3):
    for _ev in _BLOCK_EVENTS:
        _sl.bind(_ev, _block_mouse)
    # Also block on the internal canvas that customtkinter uses
    if hasattr(_sl, "_canvas"):
        for _ev in _BLOCK_EVENTS:
            _sl._canvas.bind(_ev, _block_mouse)

# --- VU METERS (placed to the right of each slider, 4px gap) ---
vu1 = VUMeter(root, width=8, height=400)
vu1.place(x=337, y=150)

vu2 = VUMeter(root, width=8, height=400)
vu2.place(x=447, y=150)

vu3 = VUMeter(root, width=8, height=400)
vu3.place(x=557, y=150)

_vu_peak = [0.0, 0.0, 0.0]   # peak-hold values per channel
_VU_PEAK_DECAY = 0.012        # amount peak drops per 33ms frame
_VU_DB_FLOOR   = -30.0        # dB level that maps to 0% on the meter

def _linear_to_vu(linear: float) -> float:
    """Convert linear peak (0–1) to display level (0–1) on a dB scale.
    -50 dB → 0.0 (bottom), 0 dB → 1.0 (top). Spreads the useful range
    so apps like Discord don't permanently max out the meter."""
    if linear <= 0:
        return 0.0
    db = 20.0 * math.log10(linear)
    return max(0.0, min(1.0, (db - _VU_DB_FLOOR) / (-_VU_DB_FLOOR)))

def _draw_vu_meters():
    for i, vu in enumerate((vu1, vu2, vu3)):
        raw   = _get_channel_meter_level(i)
        raw  *= smooth_vals[i] / 1023.0    # scale by current slider position
        level = _linear_to_vu(raw)
        if level > _vu_peak[i]:
            _vu_peak[i] = level
        else:
            _vu_peak[i] = max(0.0, _vu_peak[i] - _VU_PEAK_DECAY)
        vu.set_level(level, _vu_peak[i])

def poll_vu():
    _draw_vu_meters()
    root.after(33, poll_vu)

# --- CHANNEL CONFIG ---
CHANNEL_NAMES = ["Canal 1", "Canal 2", "Canal 3", "Dial L", "Dial R"]
DEFAULTS = [
    ("master", "Volumen General"),
    ("focus",  "App en Focus"),
    ("mic",    "Microfono"),
]
_active_profile_data = config.get("profiles", {}).get(
    config.get("active_profile", "Default"), {})
_saved_assignments = _active_profile_data.get("channel_assignments", {})
channel_assignments = {i: _saved_assignments.get(str(i)) for i in range(5)}
channel_buttons = []

def save_channel_assignments():
    cfg = load_config()
    active = cfg.get("active_profile", "Default")
    cfg.setdefault("profiles", {}).setdefault(active, {})
    cfg["profiles"][active]["channel_assignments"] = {str(k): v for k, v in channel_assignments.items()}
    save_config(cfg)

# System/background processes that should never appear in the audio app list
_SYSTEM_APP_BLACKLIST = {
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

def strip_exe(name):
    return name[:-4] if name and name.lower().endswith(".exe") else name

def _is_system_app(name: str) -> bool:
    return name.lower() in _SYSTEM_APP_BLACKLIST

# --- AUDIO SESSION CACHE ---
# GetAllSessions() scans all Windows audio processes — expensive if called 25×/s.
# Cache maps app_name → LIST of interfaces (apps like Discord have multiple sessions).
_session_cache: dict = {}   # {app_name: [ISimpleAudioVolume, ...]}
_meter_cache:   dict = {}   # {app_name: [IAudioMeterInformation, ...]}
_cache_ts: float = 0.0
_CACHE_TTL = 1.5            # seconds between full rescans
_master_meter = None        # cached IAudioMeterInformation for master output
_mic_meter    = None        # cached IAudioMeterInformation for microphone

# --- APP ICON CACHE ---
_icon_cache: dict = {}      # {app_name: CTkImage | None}
_icon_cache_is_letter: set = set()  # apps cuyo icono es letra (pendiente de reintento)

class _SHFILEINFOW(ctypes.Structure):
    _fields_ = [("hIcon", ctypes.c_void_p), ("iIcon", ctypes.c_int),
                ("dwAttributes", ctypes.c_uint32),
                ("szDisplayName", ctypes.c_wchar * 260),
                ("szTypeName",    ctypes.c_wchar * 80)]

class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_int32),
                ("biHeight", ctypes.c_int32), ("biPlanes", ctypes.c_uint16),
                ("biBitCount", ctypes.c_uint16), ("biCompression", ctypes.c_uint32),
                ("biSizeImage", ctypes.c_uint32), ("biXPelsPM", ctypes.c_int32),
                ("biYPelsPM", ctypes.c_int32), ("biClrUsed", ctypes.c_uint32),
                ("biClrImportant", ctypes.c_uint32)]

def _hicon_to_pil(hicon, size: int):
    hdc_src = ctypes.windll.user32.GetDC(0)
    hdc     = ctypes.windll.gdi32.CreateCompatibleDC(hdc_src)
    hbm     = ctypes.windll.gdi32.CreateCompatibleBitmap(hdc_src, size, size)
    ctypes.windll.gdi32.SelectObject(hdc, hbm)
    ctypes.windll.gdi32.BitBlt(hdc, 0, 0, size, size, 0, 0, 0, 0x00000042)  # BLACKNESS
    ctypes.windll.user32.DrawIconEx(hdc, 0, 0, ctypes.c_void_p(hicon), size, size, 0, 0, 3)
    bih = _BITMAPINFOHEADER()
    bih.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
    bih.biWidth = size; bih.biHeight = -size
    bih.biPlanes = 1; bih.biBitCount = 32; bih.biCompression = 0
    buf = (ctypes.c_char * (size * size * 4))()
    ctypes.windll.gdi32.GetDIBits(hdc, hbm, 0, size, buf, ctypes.byref(bih), 0)
    img = Image.frombuffer("RGBA", (size, size), bytes(buf), "raw", "BGRA", 0, 1)
    ctypes.windll.gdi32.DeleteObject(hbm)
    ctypes.windll.gdi32.DeleteDC(hdc)
    ctypes.windll.user32.ReleaseDC(0, hdc_src)
    return img

def _letter_icon(letter: str, size: int = 18):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((1, 1, size - 2, size - 2), fill="#5A5A5A")
    d.text((size // 2, size // 2), letter[:1].upper(), fill="white", anchor="mm")
    return img

def _get_app_ctk_icon(app_name: str, size: int = 18):
    """Return a CTkImage for app_name. Uses Windows icon extraction, falls back to letter."""
    if app_name in _icon_cache:
        return _icon_cache[app_name]
    img = None
    exe_path = None
    # Buscar proceso corriendo
    try:
        for proc in psutil.process_iter(['name', 'exe']):
            try:
                if strip_exe(proc.info.get('name', '')).lower() == app_name.lower() and proc.info.get('exe'):
                    exe_path = proc.info['exe']
                    break
            except Exception:
                continue
    except Exception:
        pass
    # Si no está corriendo, usar path guardado en config
    if not exe_path:
        exe_path = config.get("exe_paths", {}).get(app_name.lower())
    # Extraer ícono del exe
    if exe_path and os.path.exists(exe_path):
        try:
            sfi = _SHFILEINFOW()
            if ctypes.windll.shell32.SHGetFileInfoW(
                    exe_path, 0, ctypes.byref(sfi),
                    ctypes.sizeof(sfi), 0x100 | 0x1):
                img = _hicon_to_pil(sfi.hIcon, size)
                ctypes.windll.user32.DestroyIcon(ctypes.c_void_p(sfi.hIcon))
            # Guardar path en config para próximas veces
            if "exe_paths" not in config:
                config["exe_paths"] = {}
            config["exe_paths"][app_name.lower()] = exe_path
            save_config(config)
        except Exception:
            pass
    if img is None:
        img = _letter_icon(app_name, size)
        _icon_cache_is_letter.add(app_name)
    else:
        _icon_cache_is_letter.discard(app_name)
    ctk_img = customtkinter.CTkImage(light_image=img, dark_image=img, size=(size, size))
    _icon_cache[app_name] = ctk_img
    return ctk_img


def _retry_letter_icons():
    """Reintenta cargar íconos reales para apps que quedaron con letra."""
    for app_name in list(_icon_cache_is_letter):
        _icon_cache.pop(app_name, None)
        _get_app_ctk_icon(app_name)
    # Programar próximo reintento si quedan letras pendientes
    if _icon_cache_is_letter:
        root.after(10000, _retry_letter_icons)

def _refresh_session_cache():
    global _session_cache, _meter_cache, _cache_ts, _master_meter, _mic_meter
    new_vol, new_meter = {}, {}
    try:
        for s in AudioUtilities.GetAllSessions():
            if s.Process:
                raw  = strip_exe(s.Process.name())
                if not raw or _is_system_app(raw):
                    continue
                name = raw.capitalize()   # "opera" → "Opera"
                if name not in new_vol:
                    new_vol[name]   = []
                    new_meter[name] = []
                try:
                    new_vol[name].append(s._ctl.QueryInterface(ISimpleAudioVolume))
                except Exception:
                    pass
                try:
                    new_meter[name].append(s._ctl.QueryInterface(IAudioMeterInformation))
                except Exception:
                    pass
    except Exception as e:
        print(f"[DEBUG] Cache refresh error: {e}")
        return
    _session_cache = new_vol
    _meter_cache   = new_meter
    _cache_ts      = time.monotonic()
    # (Re)cache master and mic meter interfaces
    try:
        spk = AudioUtilities.GetSpeakers()
        _master_meter = spk.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None
                        ).QueryInterface(IAudioMeterInformation)
    except Exception:
        pass
    try:
        mic = AudioUtilities.GetMicrophone()
        if mic:
            _mic_meter = mic.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None
                         ).QueryInterface(IAudioMeterInformation)
    except Exception:
        pass

def _invalidate_cache():
    global _cache_ts
    _cache_ts = 0.0  # force refresh on next access

def _get_channel_meter_level(channel_idx: int) -> float:
    """Return peak audio level (0.0–1.0) for the given channel index."""
    assignment = channel_assignments.get(channel_idx)
    if not assignment:
        return 0.0
    try:
        if assignment == "master":
            return _master_meter.GetPeakValue() if _master_meter else 0.0
        if assignment == "mic":
            return _mic_meter.GetPeakValue() if _mic_meter else 0.0
        if assignment == "focus":
            exe_lower = _last_focus_exe.lower()
            m = next((v for k, v in _meter_cache.items() if k.lower() == exe_lower), None)
            return m.GetPeakValue() if m else 0.0
        meters = _meter_cache.get(assignment, [])
        peak, failed = 0.0, False
        for m in meters:
            try:
                peak = max(peak, m.GetPeakValue())
            except Exception:
                failed = True
        if failed:
            _invalidate_cache()
        return peak
    except Exception:
        return 0.0

def _cache_refresh_loop():
    """Background thread: refresh audio session cache every _CACHE_TTL seconds."""
    comtypes.CoInitialize()
    while True:
        try:
            _refresh_session_cache()
        except Exception as e:
            print(f"[DEBUG] Cache refresh loop error: {e}")
        time.sleep(_CACHE_TTL)

threading.Thread(target=_cache_refresh_loop, daemon=True).start()

def _get_vol_ctl(app_name):
    return _session_cache.get(app_name, [])

def get_audio_apps():
    _refresh_session_cache()   # always fresh list when user opens config window
    apps = list(_session_cache.keys())
    print(f"[DEBUG] Audio apps: {apps}")
    return apps

def set_master_volume(level):
    AudioUtilities.GetSpeakers().EndpointVolume.SetMasterVolumeLevelScalar(level, None)

def set_mic_volume(level):
    mic = AudioUtilities.GetMicrophone()
    if mic:
        interface = mic.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        volume.SetMasterVolumeLevelScalar(level, None)

def set_app_volume(display_name, level):
    ctls = _get_vol_ctl(display_name)
    failed = False
    for ctl in ctls:
        try:
            ctl.SetMasterVolume(level, None)
        except Exception:
            failed = True
    if failed:
        _invalidate_cache()   # session died — rescan next tick

def set_focus_app_volume(level, raw_val=None):
    global _focus_prev_raw
    exe = _last_focus_exe
    if not exe:
        _focus_prev_raw = raw_val
        return
    # Only claim priority if the focus slider is actually moving
    focus_moving = raw_val is not None and (
        _focus_prev_raw is None or abs(raw_val - _focus_prev_raw) > 15
    )
    _focus_prev_raw = raw_val
    if focus_moving and _channel_priority.get(exe) != "focus":
        # Record where the direct slider was so we can detect when user moves it back
        for ch, app in channel_assignments.items():
            if app == exe and _last_raw_per_channel[ch] is not None:
                _focus_took_at[exe] = _last_raw_per_channel[ch]
                break
        _channel_priority[exe] = "focus"
    if _channel_priority.get(exe, "direct") == "focus":
        set_app_volume(exe, level)

def apply_volume(channel_idx, raw_val):
    assignment = channel_assignments.get(channel_idx)
    if not assignment:
        return
    level = raw_to_level(raw_val)
    try:
        if assignment == "master":
            set_master_volume(level)
        elif assignment == "mic":
            set_mic_volume(level)
        elif assignment == "focus":
            set_focus_app_volume(level, raw_val)
        else:
            _last_raw_per_channel[channel_idx] = raw_val
            if _channel_priority.get(assignment, "direct") == "focus":
                # Reclaim direct priority if slider moved >40 units from when focus took over
                ref = _focus_took_at.get(assignment)
                if ref is not None and abs(raw_val - ref) > 40:
                    _channel_priority[assignment] = "direct"
                    _focus_took_at.pop(assignment, None)
            if _channel_priority.get(assignment, "direct") == "direct":
                set_app_volume(assignment, level)
    except Exception as e:
        print(f"[DEBUG] Volume ch{channel_idx}: {e}")

# --- ENCODER BUTTON CONFIG ---
# Actions for click / double-click
ENCODER_CLICK_ACTIONS = [
    ("none",       "Ninguna"),
    ("mute",       "Mute / Unmute"),
    ("play_pause", "Play / Pausa"),
    ("next_track", "Siguiente pista"),
    ("prev_track", "Pista anterior"),
    ("vol_up",     "Vol. Maestro +5%"),
    ("vol_down",   "Vol. Maestro -5%"),
    ("hotkey",     "Atajo de teclado"),
]
# Actions for click+turn (which channel to control while held)
ENCODER_TURN_CHANNELS = [
    ("none",          "Canal asignado (normal)"),
    ("master",        "Volumen General"),
    ("mic",           "Microfono"),
    ("focus",         "App en Focus"),
    ("output_device", "Cambiar salida de audio"),
]
ENCODER_NAMES = ["Enc. Izq.", "Enc. Der."]
ACTION_TYPES = [
    ("click",      "Click"),
    ("dblclick",   "Doble Click"),
    ("click_turn", "Click + Girar"),
]

_saved_enc = _active_profile_data.get("encoder_assignments", [{}, {}])
encoder_assignments = []
for _i in range(2):
    _d = _saved_enc[_i] if _i < len(_saved_enc) else {}
    _entry = {atype: _d.get(atype, "none") for atype, _ in ACTION_TYPES}
    _entry["hotkey_keys_click"]    = _d.get("hotkey_keys_click", "")
    _entry["hotkey_keys_dblclick"] = _d.get("hotkey_keys_dblclick", "")
    encoder_assignments.append(_entry)
encoder_held = [False, False]
enc_config_buttons = []

def save_encoder_assignments():
    cfg = load_config()
    active = cfg.get("active_profile", "Default")
    cfg.setdefault("profiles", {}).setdefault(active, {})
    cfg["profiles"][active]["encoder_assignments"] = [dict(enc) for enc in encoder_assignments]
    save_config(cfg)

# Windows media key VK codes
_VK_MEDIA_PLAY_PAUSE = 0xB3
_VK_MEDIA_NEXT_TRACK = 0xB0
_VK_MEDIA_PREV_TRACK = 0xB1
_KEYEVENTF_EXTENDEDKEY = 0x0001
_KEYEVENTF_KEYUP       = 0x0002

def _send_media_key(vk):
    # Media keys require KEYEVENTF_EXTENDEDKEY on Windows
    ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY, 0)
    ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP, 0)

_VK_MAP = {
    'ctrl': 0x11, 'shift': 0x10, 'alt': 0x12, 'win': 0x5B,
    'space': 0x20, 'enter': 0x0D, 'return': 0x0D, 'esc': 0x1B, 'escape': 0x1B,
    'tab': 0x09, 'backspace': 0x08, 'delete': 0x2E, 'insert': 0x2D,
    'home': 0x24, 'end': 0x23, 'pageup': 0x21, 'prior': 0x21,
    'pagedown': 0x22, 'next': 0x22,
    'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74,
    'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79,
    'f11': 0x7A, 'f12': 0x7B,
}

def _send_hotkey(hotkey_str):
    if not hotkey_str:
        return
    parts = [p.strip().lower() for p in hotkey_str.split('+')]
    MOD_ORDER = ['ctrl', 'shift', 'alt', 'win']
    modifiers = [p for p in parts if p in MOD_ORDER]
    main_keys = [p for p in parts if p not in MOD_ORDER]
    if not main_keys:
        return
    def get_vk(key):
        if key in _VK_MAP:
            return _VK_MAP[key]
        if len(key) == 1:
            return ctypes.windll.user32.VkKeyScanW(ord(key)) & 0xFF
        return None
    try:
        for mod in modifiers:
            ctypes.windll.user32.keybd_event(_VK_MAP[mod], 0, 0, 0)
        for k in main_keys:
            vk = get_vk(k)
            if vk:
                ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY, 0)
                ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP, 0)
        for mod in reversed(modifiers):
            ctypes.windll.user32.keybd_event(_VK_MAP[mod], 0, _KEYEVENTF_KEYUP, 0)
    except Exception as e:
        print(f"[DEBUG] send_hotkey error: {e}")

def _toggle_mute_by_assignment(assignment):
    try:
        if assignment == "master":
            ep = AudioUtilities.GetSpeakers().EndpointVolume
            ep.SetMute(not ep.GetMute(), None)
        elif assignment == "mic":
            mic = AudioUtilities.GetMicrophone()
            if mic:
                interface = mic.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                vol = interface.QueryInterface(IAudioEndpointVolume)
                vol.SetMute(not vol.GetMute(), None)
        elif assignment == "focus":
            _mute_app(_last_focus_exe)
        else:
            _mute_app(assignment)
    except Exception as e:
        print(f"[DEBUG] Mute toggle error: {e}")

def _mute_app(app_name):
    for s in AudioUtilities.GetAllSessions():
        if s.Process and strip_exe(s.Process.name()) == app_name:
            vol = s._ctl.QueryInterface(ISimpleAudioVolume)
            vol.SetMute(not vol.GetMute(), None)

def execute_encoder_action(enc_idx, action_type):
    action = encoder_assignments[enc_idx].get(action_type, "none")
    print(f"[BTN] execute enc={enc_idx} type={action_type} action={action}")
    if action == "none":
        return
    try:
        if action == "mute":
            assignment = channel_assignments.get(3 + enc_idx)
            if assignment:
                if not _dial_muted[enc_idx]:
                    # Mute: save position, snap dial to 0, freeze physical input
                    _dial_saved_val[enc_idx] = max(smooth_vals[3 + enc_idx], 16.0)
                    smooth_vals[3 + enc_idx] = 0.0
                    target_vals[3 + enc_idx] = 0.0
                    _dial_muted[enc_idx] = True
                    root.after(0, update_channel_buttons)
                    root.after(0, _apply_ui_from_smooth)
                    threading.Thread(
                        target=_apply_volume_by_key, args=(assignment, 0), daemon=True
                    ).start()
                else:
                    # Unmute: restore saved position
                    _dial_muted[enc_idx] = False
                    saved = _dial_saved_val[enc_idx]
                    smooth_vals[3 + enc_idx] = saved
                    target_vals[3 + enc_idx] = saved
                    root.after(0, update_channel_buttons)
                    root.after(0, _apply_ui_from_smooth)
                    threading.Thread(
                        target=_apply_volume_by_key, args=(assignment, saved), daemon=True
                    ).start()
        elif action == "play_pause":
            _send_media_key(_VK_MEDIA_PLAY_PAUSE)
        elif action == "next_track":
            _send_media_key(_VK_MEDIA_NEXT_TRACK)
        elif action == "prev_track":
            _send_media_key(_VK_MEDIA_PREV_TRACK)
        elif action == "vol_up":
            ep = AudioUtilities.GetSpeakers().EndpointVolume
            set_master_volume(min(1.0, ep.GetMasterVolumeLevelScalar() + 0.05))
        elif action == "vol_down":
            ep = AudioUtilities.GetSpeakers().EndpointVolume
            set_master_volume(max(0.0, ep.GetMasterVolumeLevelScalar() - 0.05))
        elif action == "hotkey":
            hk_str = encoder_assignments[enc_idx].get(f"hotkey_keys_{action_type}", "")
            if hk_str:
                threading.Thread(target=_send_hotkey, args=(hk_str,), daemon=True).start()
    except Exception as e:
        print(f"[DEBUG] Encoder action error: {e}")

def handle_button_event(enc_idx, event_code):
    """event_code: 1=click, 2=dblclick, 3=held_start, 4=held_end"""
    names = {1: "click", 2: "dblclick", 3: "held_start", 4: "held_end"}
    print(f"[BTN] enc={enc_idx} event={names.get(event_code, event_code)}")
    if event_code == 1:
        execute_encoder_action(enc_idx, "click")
    elif event_code == 2:
        execute_encoder_action(enc_idx, "dblclick")
    elif event_code == 3:
        encoder_held[enc_idx] = True
        turn_key = encoder_assignments[enc_idx].get("click_turn", "none")
        if turn_key != "none":
            _enc_hold_saved[enc_idx] = smooth_vals[3 + enc_idx]
    elif event_code == 4:
        encoder_held[enc_idx] = False
        turn_key = encoder_assignments[enc_idx].get("click_turn", "none")
        if turn_key != "none" and _enc_hold_saved[enc_idx] is not None:
            smooth_vals[3 + enc_idx] = _enc_hold_saved[enc_idx]
            target_vals[3 + enc_idx] = _enc_hold_saved[enc_idx]
        _enc_hold_saved[enc_idx] = None
        _enc_prev_raw[enc_idx] = None

def open_encoder_config(enc_idx):
    t = THEMES[current_theme]
    win = customtkinter.CTkToplevel(root)
    win.title(f"Configurar {ENCODER_NAMES[enc_idx]}")
    win.geometry("310x280")
    win.resizable(False, False)
    win.grab_set()
    win.lift()
    win.attributes('-topmost', True)
    win.after(150, lambda: win.attributes('-topmost', False))
    win.configure(fg_color=t["bg"])

    customtkinter.CTkLabel(
        win, text=f"Acciones — {ENCODER_NAMES[enc_idx]}",
        font=customtkinter.CTkFont(size=14, weight="bold"),
        text_color=t["text"]
    ).pack(pady=(12, 8))

    dropdown_vars = {}
    hotkey_vars   = {}   # atype -> StringVar with recorded combo
    click_labels  = [v for _, v in ENCODER_CLICK_ACTIONS]
    turn_labels   = [v for _, v in ENCODER_TURN_CHANNELS]
    click_key_map = {v: k for k, v in ENCODER_CLICK_ACTIONS}
    turn_key_map  = {v: k for k, v in ENCODER_TURN_CHANNELS}

    MOD_KEYS  = {'Control_L','Control_R','Shift_L','Shift_R','Alt_L','Alt_R','Super_L','Super_R'}
    MOD_NAMES = {'Control_L':'ctrl','Control_R':'ctrl','Shift_L':'shift','Shift_R':'shift',
                 'Alt_L':'alt','Alt_R':'alt','Super_L':'win','Super_R':'win'}

    for atype, alabel in ACTION_TYPES:
        container = customtkinter.CTkFrame(win, fg_color="transparent")
        container.pack(fill="x")

        row = customtkinter.CTkFrame(container, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(4, 0))
        customtkinter.CTkLabel(
            row, text=alabel, width=115, anchor="w",
            text_color=t["text"], font=customtkinter.CTkFont(size=12)
        ).pack(side="left")

        if atype == "click_turn":
            opts = turn_labels
            cur_key = encoder_assignments[enc_idx].get(atype, "none")
            cur_label = next((v for k, v in ENCODER_TURN_CHANNELS if k == cur_key),
                             ENCODER_TURN_CHANNELS[0][1])
            var = customtkinter.StringVar(value=cur_label)
            dropdown_vars[atype] = var
            customtkinter.CTkOptionMenu(
                row, values=opts, variable=var, width=170,
                fg_color=t["btn_un"], button_color=t["accent"],
                button_hover_color=t["accent_hover"], text_color=t["text"],
                dropdown_fg_color=t["bg_mid"], dropdown_text_color=t["text"],
                dropdown_hover_color=t["accent"]
            ).pack(side="left")
        else:
            opts = click_labels
            cur_key = encoder_assignments[enc_idx].get(atype, "none")
            cur_label = next((v for k, v in ENCODER_CLICK_ACTIONS if k == cur_key),
                             ENCODER_CLICK_ACTIONS[0][1])
            var = customtkinter.StringVar(value=cur_label)
            dropdown_vars[atype] = var

            hk_key = f"hotkey_keys_{atype}"
            hk_str = encoder_assignments[enc_idx].get(hk_key, "")
            hk_var = customtkinter.StringVar(value=hk_str)
            hotkey_vars[atype] = hk_var

            # Hotkey recorder row (shown only when "Atajo de teclado" is selected)
            hk_row = customtkinter.CTkFrame(container, fg_color="transparent")
            btn_label = hk_str if hk_str else "Grabar atajo..."
            grabar_btn = customtkinter.CTkButton(
                hk_row, text=btn_label, width=170, height=28,
                fg_color=t["btn_un"], hover_color=t["accent_hover"],
                text_color=t["text"], font=customtkinter.CTkFont(size=11)
            )
            grabar_btn.pack(side="left", padx=(119, 0))

            def make_capture(at, hkv, gbtn):
                def start_capture():
                    gbtn.configure(text="Presioná las teclas...", fg_color=t["accent"])
                    win.focus_force()
                    pressed_mods = set()
                    def on_key(e):
                        if e.keysym in MOD_KEYS:
                            pressed_mods.add(MOD_NAMES[e.keysym])
                            return
                        win.unbind('<KeyPress>')
                        main = e.keysym.lower()
                        order = ['ctrl','shift','alt','win']
                        parts = [m for m in order if m in pressed_mods] + [main]
                        combo = '+'.join(parts)
                        hkv.set(combo)
                        gbtn.configure(text=combo, fg_color=t["btn_un"])
                    win.bind('<KeyPress>', on_key)
                return start_capture

            grabar_btn.configure(command=make_capture(atype, hk_var, grabar_btn))

            def make_dd_callback(hkrow, hkv):
                def on_change(label):
                    key = click_key_map.get(label, "none")
                    if key == "hotkey":
                        hkrow.pack(fill="x", padx=14, pady=(0, 4))
                        win.geometry(f"310x{win.winfo_height() + 38}")
                    else:
                        if hkrow.winfo_ismapped():
                            hkrow.pack_forget()
                            win.geometry(f"310x{max(280, win.winfo_height() - 38)}")
                return on_change

            customtkinter.CTkOptionMenu(
                row, values=opts, variable=var, width=170,
                command=make_dd_callback(hk_row, hk_var),
                fg_color=t["btn_un"], button_color=t["accent"],
                button_hover_color=t["accent_hover"], text_color=t["text"],
                dropdown_fg_color=t["bg_mid"], dropdown_text_color=t["text"],
                dropdown_hover_color=t["accent"]
            ).pack(side="left")

            if cur_key == "hotkey":
                hk_row.pack(fill="x", padx=14, pady=(0, 4))
                win.geometry(f"310x{280 + 38}")

    def save_and_close():
        for atype, var in dropdown_vars.items():
            if atype == "click_turn":
                encoder_assignments[enc_idx][atype] = turn_key_map.get(var.get(), "none")
            else:
                encoder_assignments[enc_idx][atype] = click_key_map.get(var.get(), "none")
                encoder_assignments[enc_idx][f"hotkey_keys_{atype}"] = hotkey_vars[atype].get()
        save_encoder_assignments()
        win.destroy()

    customtkinter.CTkButton(
        win, text="Guardar", height=34,
        fg_color=t["accent"], hover_color=t["accent_hover"], text_color=t["text"],
        command=save_and_close
    ).pack(pady=10)

def assignment_label(val):
    mapping = {"master": "Vol. General", "focus": "App en Focus",
               "mic": "Microfono", None: "Sin asignar"}
    return mapping.get(val, val) if val in mapping else val

_prev_muted = [False] * 5   # track last mute state to detect changes

def update_channel_buttons():
    t = THEMES[current_theme]
    for i, btn in enumerate(channel_buttons):
        assigned = channel_assignments.get(i)
        muted    = bool(assigned) and smooth_vals[i] <= MUTE_THRESHOLD
        _prev_muted[i] = muted

        icon = None
        if assigned and assigned not in ("master", "mic", "focus"):
            icon = _get_app_ctk_icon(assigned)

        label = assignment_label(assigned)
        if muted:
            fg_col   = t["btn_un"]
            txt_col  = t["text_dim"]
            label    = "⦸ " + label
        else:
            fg_col  = t["accent"] if assigned else t["btn_un"]
            txt_col = t["text"]

        if icon:
            btn.configure(text=label, fg_color=fg_col, hover_color=t["accent_hover"],
                          text_color=txt_col, image=icon, compound="left")
        else:
            btn.configure(text=label, fg_color=fg_col, hover_color=t["accent_hover"],
                          text_color=txt_col, image=None, compound="center")

def _make_channel_number_icon(number: int, size: int = 64) -> Image.Image:
    """Create a square PIL image with the channel number centered."""
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([2, 2, size - 3, size - 3], radius=size // 5,
                            fill="#555555")
    try:
        from PIL import ImageFont
        font = ImageFont.truetype("arialbd.ttf", int(size * 0.55))
    except Exception:
        font = None
    text = str(number)
    if font:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text(((size - tw) / 2 - bbox[0], (size - th) / 2 - bbox[1]),
                  text, fill="white", font=font)
    else:
        draw.text((size // 2, size // 2), text, fill="white", anchor="mm")
    return img

def open_channel_config(channel_idx):
    t = THEMES[current_theme]
    win = customtkinter.CTkToplevel(root)
    win.title(f"Configurar {CHANNEL_NAMES[channel_idx]}")
    win.geometry("320x480")
    win.resizable(False, False)
    win.grab_set()
    win.lift()
    win.attributes('-topmost', True)
    win.after(150, lambda: win.attributes('-topmost', False))
    win.configure(fg_color=t["bg"])

    # Set numbered icon in the title bar
    try:
        _num_pil = _make_channel_number_icon(channel_idx + 1, 64)
        _ico_tmp = os.path.join(_tempfile.gettempdir(), f"knob_ch{channel_idx+1}.ico")
        _num_pil.save(_ico_tmp, format="ICO", sizes=[(64, 64), (32, 32), (16, 16)])
        win.after(50, lambda: win.iconbitmap(_ico_tmp))
    except Exception:
        pass

    customtkinter.CTkLabel(
        win, text=f"Asignar {CHANNEL_NAMES[channel_idx]}",
        font=customtkinter.CTkFont(size=15, weight="bold"),
        text_color=t["text"]
    ).pack(pady=12)

    def assign(val):
        channel_assignments[channel_idx] = val
        save_channel_assignments()
        update_channel_buttons()
        win.destroy()

    customtkinter.CTkLabel(win, text="Opciones por defecto",
        font=customtkinter.CTkFont(size=12), text_color=t["text_dim"]).pack(anchor="w", padx=18)
    defaults_frame = customtkinter.CTkFrame(win, fg_color=t["bg_mid"])
    defaults_frame.pack(fill="x", padx=15, pady=(4, 10))
    for key, label in DEFAULTS:
        customtkinter.CTkButton(
            defaults_frame, text=label, height=32,
            fg_color=t["accent"] if channel_assignments[channel_idx] == key else t["btn_un"],
            hover_color=t["accent_hover"], text_color=t["text"],
            command=lambda k=key: assign(k)
        ).pack(fill="x", padx=10, pady=3)

    customtkinter.CTkLabel(win, text="Aplicaciones en ejecucion",
        font=customtkinter.CTkFont(size=12), text_color=t["text_dim"]).pack(anchor="w", padx=18)
    apps_frame = customtkinter.CTkScrollableFrame(win, height=220, fg_color=t["bg_mid"])
    apps_frame.pack(fill="x", padx=15, pady=(4, 10))

    apps = get_audio_apps()
    if not apps:
        customtkinter.CTkLabel(apps_frame, text="No hay apps con audio activo",
            text_color="gray60").pack(pady=10)
    else:
        for app in apps:
            customtkinter.CTkButton(
                apps_frame, text=app, height=30,
                fg_color=t["accent"] if channel_assignments[channel_idx] == app else t["btn_un"],
                hover_color=t["accent_hover"], text_color=t["text"],
                command=lambda a=app: assign(a)
            ).pack(fill="x", pady=2)

# Buttons above sliders — centered over each slider
# slider width=25, center=x+12; btn width=80, x=center-40
# slider1 center=352 → x=312; slider2 462 → 422; slider3 572 → 532
slider_btn_data = [(280, 112, 0), (390, 112, 1), (500, 112, 2)]
for bx, by, idx in slider_btn_data:
    btn = customtkinter.CTkButton(
        root, text="Sin asignar",
        width=80, height=38,
        font=customtkinter.CTkFont(size=11),
        command=lambda i=idx: open_channel_config(i)
    )
    btn.place(x=bx, y=by)
    channel_buttons.append(btn)

# Buttons above dials — centered over dial circles
# dial circle center ≈ x+65; btn width=90, x=circle_center-45
# dial_izq: circle≈180, btn x=135; dial_der: circle≈400, btn x=355
dial_btn_data = [(85, 565, 3), (310, 565, 4)]
for bx, by, idx in dial_btn_data:
    btn = customtkinter.CTkButton(
        frame_derecho, text="Sin asignar",
        width=90, height=38,
        font=customtkinter.CTkFont(size=11),
        command=lambda i=idx: open_channel_config(i)
    )
    btn.place(x=bx, y=by)
    channel_buttons.append(btn)

# Encoder config buttons flanking the right dial
# Left dial (enc 0):  button at x=215 (between dials), y=668 (vertically centered on dial)
# Right dial (enc 1): button at x=425 (right of right dial), y=668
enc_btn_positions = [(100, 750, 0), (325, 750, 1)]
for bx, by, enc_idx in enc_btn_positions:
    btn = customtkinter.CTkButton(
        frame_derecho,
        text=f"Enc {'Izq' if enc_idx == 0 else 'Der'}",
        width=65, height=30,
        font=customtkinter.CTkFont(size=10),
        command=lambda i=enc_idx: open_encoder_config(i)
    )
    btn.place(x=bx, y=by)
    enc_config_buttons.append(btn)

# --- PROFILE SYSTEM ---
_profile_menu_widget = None   # reference kept for apply_theme

def _profile_names():
    return list(load_config().get("profiles", {"Default": {}}).keys())

def _active_profile_name():
    return load_config().get("active_profile", "Default")

def switch_profile(name):
    cfg = load_config()
    p = cfg.get("profiles", {}).get(name, {})
    saved_ch  = p.get("channel_assignments", {})
    saved_enc = p.get("encoder_assignments", [{}, {}])
    for i in range(5):
        channel_assignments[i] = saved_ch.get(str(i))
    for i in range(2):
        enc = saved_enc[i] if i < len(saved_enc) else {}
        for atype, _ in ACTION_TYPES:
            encoder_assignments[i][atype] = enc.get(atype, "none")
    cfg["active_profile"] = name
    save_config(cfg)
    update_channel_buttons()
    if _profile_menu_widget:
        _profile_menu_widget.set(name)

def create_profile(name):
    """Save current assignments as a new profile and switch to it."""
    cfg = load_config()
    cfg.setdefault("profiles", {})[name] = {
        "channel_assignments": {str(k): v for k, v in channel_assignments.items()},
        "encoder_assignments": [dict(enc) for enc in encoder_assignments]
    }
    cfg["active_profile"] = name
    save_config(cfg)
    if _profile_menu_widget:
        _profile_menu_widget.configure(values=_profile_names())
        _profile_menu_widget.set(name)

def delete_active_profile():
    cfg = load_config()
    profiles = cfg.get("profiles", {})
    active = cfg.get("active_profile", "Default")
    if len(profiles) <= 1:
        return   # never delete the last profile
    del profiles[active]
    new_active = next(iter(profiles))
    cfg["active_profile"] = new_active
    save_config(cfg)
    if _profile_menu_widget:
        _profile_menu_widget.configure(values=list(profiles.keys()))
    switch_profile(new_active)

def open_new_profile_dialog():
    t = THEMES[current_theme]
    win = customtkinter.CTkToplevel(root)
    win.title("Nuevo perfil")
    win.geometry("270x155")
    win.resizable(False, False)
    win.grab_set()
    win.lift()
    win.attributes('-topmost', True)
    win.after(150, lambda: win.attributes('-topmost', False))
    win.configure(fg_color=t["bg"])

    customtkinter.CTkLabel(
        win, text="Nombre del perfil",
        font=customtkinter.CTkFont(size=13, weight="bold"),
        text_color=t["text"]
    ).pack(pady=(16, 6))

    entry = customtkinter.CTkEntry(
        win, width=210, placeholder_text="ej. Gaming, Trabajo...",
        fg_color=t["bg_mid"], text_color=t["text"], border_color=t["accent"]
    )
    entry.pack(pady=4)
    entry.focus()

    def confirm():
        name = entry.get().strip()
        if not name:
            return
        create_profile(name)
        win.destroy()

    entry.bind("<Return>", lambda e: confirm())
    btn_row = customtkinter.CTkFrame(win, fg_color="transparent")
    btn_row.pack(pady=10)
    customtkinter.CTkButton(
        btn_row, text="Guardar", width=95, height=32,
        fg_color=t["accent"], hover_color=t["accent_hover"], text_color=t["text"],
        command=confirm
    ).pack(side="left", padx=4)
    customtkinter.CTkButton(
        btn_row, text="Cancelar", width=95, height=32,
        fg_color=t["btn_un"], hover_color=t["accent_hover"], text_color=t["text"],
        command=win.destroy
    ).pack(side="left", padx=4)

# --- STARTUP ---
APP_REG_NAME = "KNOBApp"
if getattr(sys, 'frozen', False):
    APP_PATH = f'"{sys.executable}" --minimized'
else:
    APP_PATH = f'"{sys.executable}" "{os.path.abspath(__file__)}" --minimized'

def _refresh_startup_path():
    """Si el autostart estaba activado, actualiza el path al ejecutable actual."""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_READ | winreg.KEY_SET_VALUE)
        try:
            winreg.QueryValueEx(key, APP_REG_NAME)
            winreg.SetValueEx(key, APP_REG_NAME, 0, winreg.REG_SZ, APP_PATH)
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
    except Exception:
        pass
_refresh_startup_path()

def is_startup_enabled():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_READ)
        winreg.QueryValueEx(key, APP_REG_NAME)
        winreg.CloseKey(key)
        return True
    except Exception:
        return False

def toggle_startup():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_SET_VALUE)
        if var_startup.get():
            winreg.SetValueEx(key, APP_REG_NAME, 0, winreg.REG_SZ, APP_PATH)
        else:
            try:
                winreg.DeleteValue(key, APP_REG_NAME)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception as e:
        print(f"[DEBUG] Startup registry error: {e}")

# --- TRAY ---
tray_icon_obj = None

def show_from_tray(icon, _item):
    global tray_icon_obj
    icon.stop()
    tray_icon_obj = None
    root.after(0, root.deiconify)

def quit_from_tray(icon, _item):
    global tray_icon_obj
    icon.stop()
    tray_icon_obj = None
    root.after(0, root.destroy)

def hide_to_tray():
    global tray_icon_obj
    root.withdraw()
    img = _app_icon_pil.copy().convert("RGB").resize((64, 64), Image.LANCZOS)
    menu = pystray.Menu(
        pystray.MenuItem("Mostrar KNOB", show_from_tray, default=True),
        pystray.MenuItem("Salir", quit_from_tray)
    )
    tray_icon_obj = pystray.Icon("KNOB", img, "KNOB App", menu)
    threading.Thread(target=tray_icon_obj.run, daemon=True).start()

def on_close():
    global ser_global, conectado
    if var_tray.get() and PYSTRAY_AVAILABLE:
        hide_to_tray()   # just hide — keep serial running in background
        return
    # Full exit: stop serial thread then destroy window
    conectado = False
    if ser_global:
        try:
            ser_global.close()
        except Exception:
            pass
        ser_global = None
    root.destroy()

# --- CONFIG TAB CONTENT ---
var_startup = customtkinter.BooleanVar(value=is_startup_enabled())
var_tray = customtkinter.BooleanVar(value=config.get("minimize_to_tray", False))

# Profiles section
lbl_perfiles = customtkinter.CTkLabel(content_config, text="Perfiles",
    font=customtkinter.CTkFont(size=12, weight="bold"))
lbl_perfiles.pack(anchor="w", padx=10, pady=(12, 2))

_profile_menu_widget = customtkinter.CTkOptionMenu(
    content_config,
    values=_profile_names(),
    command=switch_profile,
    width=118, height=30,
    font=customtkinter.CTkFont(size=11),
    dynamic_resizing=False
)
_profile_menu_widget.set(_active_profile_name())
_profile_menu_widget.pack(fill="x", padx=10, pady=(0, 4))

_profile_btn_row = customtkinter.CTkFrame(content_config, fg_color="transparent")
_profile_btn_row.pack(fill="x", padx=10, pady=(0, 6))

_btn_nuevo = customtkinter.CTkButton(
    _profile_btn_row, text="Nuevo", height=28, width=55,
    font=customtkinter.CTkFont(size=11),
    command=open_new_profile_dialog
)
_btn_nuevo.pack(side="left", padx=(0, 4))

_btn_borrar = customtkinter.CTkButton(
    _profile_btn_row, text="Borrar", height=28, width=55,
    font=customtkinter.CTkFont(size=11),
    command=delete_active_profile
)
_btn_borrar.pack(side="left")

lbl_sistema = customtkinter.CTkLabel(content_config, text="Sistema",
    font=customtkinter.CTkFont(size=12, weight="bold"))
lbl_sistema.pack(anchor="w", padx=10, pady=(12, 2))

chk_startup = customtkinter.CTkCheckBox(
    content_config,
    text="Iniciar con\nWindows",
    variable=var_startup,
    command=toggle_startup,
    font=customtkinter.CTkFont(size=11)
)
chk_startup.pack(anchor="w", padx=10, pady=4)

def toggle_tray():
    cfg = load_config()
    cfg["minimize_to_tray"] = var_tray.get()
    save_config(cfg)

chk_tray = customtkinter.CTkCheckBox(
    content_config,
    text="Cerrar al tray",
    variable=var_tray,
    command=toggle_tray,
    font=customtkinter.CTkFont(size=11),
    state="normal" if PYSTRAY_AVAILABLE else "disabled"
)
chk_tray.pack(anchor="w", padx=10, pady=4)

# Theme section
lbl_tema = customtkinter.CTkLabel(content_config, text="Tema",
    font=customtkinter.CTkFont(size=12, weight="bold"))
lbl_tema.pack(anchor="w", padx=10, pady=(12, 4))

swatch_frame = customtkinter.CTkFrame(content_config, fg_color="transparent")
swatch_frame.pack(fill="x", padx=6, pady=2)

THEME_LIST = [
    ("Magma",   "#B31D1D"),
    ("Cobalt",   "#2533B0"),
    ("Onyx",  "#111111"),  # show slightly lighter so it's visible
    ("Arctic", "#FFFFFF"),
    ("Steel",   "#B6B5B5"),
    ("Nebula", "#621694"),
]

swatch_buttons = {}
for i, (tname, tcolor) in enumerate(THEME_LIST):
    row, col = i // 2, i % 2
    txt_color = "#111111" if tname in ("Arctic", "Steel") else "#FFFFFF"
    btn = customtkinter.CTkButton(
        swatch_frame,
        text=tname,
        width=55, height=32,
        fg_color=tcolor,
        hover_color=THEMES[tname]["accent_hover"],
        text_color=txt_color,
        border_width=1,
        border_color="#444444",
        font=customtkinter.CTkFont(size=10, weight="bold"),
        corner_radius=6,
        command=lambda n=tname: apply_theme(n)
    )
    btn.grid(row=row, column=col, padx=3, pady=3)
    swatch_buttons[tname] = btn

# --- APPLY THEME ---
def apply_theme(name):
    global current_theme
    current_theme = name
    t = THEMES[name]
    sel_border = "#000000" if name == "Arctic" else "#FFFFFF"

    # Root + frames
    root.configure(fg_color=t["bg"])
    left_panel.configure(fg_color=t["bg_mid"])
    frame_derecho.configure(fg_color=t["bg"])

    # KNOB title + icon bg
    label2.configure(text_color=t["text"])
    icono_estado.configure(bg_color=t["bg"])

    # Tab buttons
    _update_tab_buttons()

    # Channel buttons (text + color)
    update_channel_buttons()

    # Sliders
    for sl in (slider1, slider2, slider3):
        sl.configure(
            button_color=t["accent"],
            button_hover_color=t["accent_hover"],
            progress_color=t["accent"],
            fg_color=t["btn_un"],
            bg_color=t["bg"]
        )

    # Dials bg + needle color (needle must contrast against bg)
    for dial in (knob_izquierda, knob_derecha):
        dial.configure(bg=t["bg"])
        dial._needle_color = t["text"]
        if dial._needle_id:
            dial.itemconfig(dial._needle_id, fill=t["text"])

    # VU meter segment separator color = bg (so gaps look transparent)
    for vu in (vu1, vu2, vu3):
        vu.update_bg(t["bg"])

    # Config tab labels
    lbl_perfiles.configure(text_color=t["text"])
    lbl_sistema.configure(text_color=t["text"])
    lbl_tema.configure(text_color=t["text"])
    lbl_status_knob.configure(
        text_color="#00C853" if conectado else t["text_dim"]
    )
    lbl_audio_device.configure(text_color=t["text_dim"])

    # Profile widgets
    _profile_menu_widget.configure(
        fg_color=t["btn_un"], button_color=t["accent"],
        button_hover_color=t["accent_hover"], text_color=t["text"],
        dropdown_fg_color=t["bg_mid"], dropdown_text_color=t["text"],
        dropdown_hover_color=t["accent"]
    )
    for btn in (_btn_nuevo, _btn_borrar):
        btn.configure(
            fg_color=t["btn_un"], hover_color=t["accent_hover"], text_color=t["text"]
        )

    # Encoder config buttons
    for btn in enc_config_buttons:
        btn.configure(
            fg_color=t["btn_un"],
            hover_color=t["accent_hover"],
            text_color=t["text"]
        )

    # Checkboxes
    for chk in (chk_startup, chk_tray):
        chk.configure(
            text_color=t["text"],
            fg_color=t["accent"],
            hover_color=t["accent_hover"],
            checkmark_color=t["text"]
        )

    # Version label / update button
    if btn_version.cget("text") == APP_VERSION:
        btn_version.configure(text_color=t["text_dim"], fg_color=t["bg_mid"])
    else:
        btn_version.configure(fg_color=t["accent"], text_color=t["text"])

    # Swatch selected state
    for n, btn in swatch_buttons.items():
        if n == name:
            btn.configure(border_width=2, border_color=sel_border)
        else:
            btn.configure(border_width=1, border_color="#444444")

    # Persist
    cfg = load_config()
    cfg["theme"] = name
    save_config(cfg)

root.protocol("WM_DELETE_WINDOW", on_close)
apply_theme(current_theme)

# Cargar dispositivos de salida de audio
threading.Thread(target=_load_output_devices, daemon=True).start()

poll_vu()

if "--minimized" in sys.argv and PYSTRAY_AVAILABLE:
    root.after(300, hide_to_tray)

threading.Thread(target=_check_for_update, daemon=True).start()

# Reintenta cargar íconos reales 10s después del arranque (apps pueden no estar corriendo aún)
root.after(10000, _retry_letter_icons)

root.mainloop()
