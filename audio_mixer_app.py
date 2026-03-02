# -*- coding: utf-8 -*-
import customtkinter as ctk
from custom_widgets import WideHandleSlider
import psutil
import os
import sys
import serial
import serial.tools.list_ports
import threading
import time
import queue
import math
import json
import traceback
import tkinter as tk
from tkinter import messagebox
import tkinter.font as tkfont

# ---- Windows single-instance / events / tray ----
import ctypes
from ctypes import wintypes

try:
    import winshell
    from win32com.client import Dispatch
except ImportError:
    winshell = None
    Dispatch = None

try:
    import win32gui
    import win32process
    import win32con
except ImportError:
    win32gui = None
    win32process = None
    win32con = None

try:
    from PIL import Image, ImageDraw, ImageOps, ImageTk
    import pystray
except ImportError:
    pystray = None
    Image = None
    ImageTk = None

from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import comtypes

# ========= Constantes de protocolo / serial =========
BAUD_RATE = 9600
INIT_TOKEN_AWAIT = "{KNOB_CMD_AWAIT_INIT}"
INIT_TOKEN_START = "{KNOB_CMD_START}\n"
INIT_TOKEN_OK = "{KNOB_LOG_INITIALIZED}"
FRAME_SEP = '|'
FRAME_VALUES = 5
SERIAL_READ_TIMEOUT = 2.0
HANDSHAKE_WINDOW_S = 3.0
NO_DATA_RECONNECT_S = 5.0

# VIDs/PIDs comunes (detectar Arduino/CH340/CP210x)
KNOWN_DEVICES = {
    (0x2341, 0x0043), (0x2341, 0x0001), (0x2341, 0x0010), (0x2341, 0x8036),
    (0x1A86, 0x7523), (0x1A86, 0x5523), (0x10C4, 0xEA60),
}

# ======== Identidad de la app ========
APP_TITLE = "KNOB"
APP_USER_MODEL_ID = "KNOB.AudioMixerApp.1.0"

# Instancia única y señal “mostrar”
MUTEX_NAME = r"Global\KNOB_AudioMixerApp_SingleInstance"
SHOW_EVENT_NAME = r"Global\KNOB_AudioMixerApp_Show"

kernel32 = ctypes.windll.kernel32

# ======== Fuentes privadas (TTF embebida) =========
FR_PRIVATE = 0x10

def load_private_font(ttf_path: str) -> bool:
    """Registra la fuente TTF en memoria (solo para este proceso)."""
    try:
        AddFontResourceEx = ctypes.windll.gdi32.AddFontResourceExW
        AddFontResourceEx.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPVOID]
        AddFontResourceEx.restype = wintypes.INT
        return AddFontResourceEx(ttf_path, FR_PRIVATE, None) > 0
    except Exception:
        return False

def unload_private_font(ttf_path: str):
    """Quita la fuente registrada en memoria al cerrar."""
    try:
        RemoveFontResourceEx = ctypes.windll.gdi32.RemoveFontResourceExW
        RemoveFontResourceEx.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPVOID]
        RemoveFontResourceEx.restype = wintypes.BOOL
        RemoveFontResourceEx(ttf_path, FR_PRIVATE, None)
    except Exception:
        pass

# ========= Utilidades =========
def resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative_path)

def set_windows_app_user_model_id(app_id: str = APP_USER_MODEL_ID):
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass

def _norm_proc_name(name: str) -> str:
    if not name:
        return ""
    n = name.strip().lower()
    if n.endswith(".exe"):
        n = n[:-4]
    return n

def _names_match(a: str, b: str) -> bool:
    return _norm_proc_name(a) == _norm_proc_name(b)

# ========= Instancia única =========
def acquire_single_instance_mutex():
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    handle = kernel32.CreateMutexW(None, True, MUTEX_NAME)
    if not handle:
        return None, True
    ERROR_ALREADY_EXISTS = 183
    already = (kernel32.GetLastError() == ERROR_ALREADY_EXISTS)
    return handle, (not already)

def release_single_instance_mutex(handle):
    try:
        if handle:
            kernel32.ReleaseMutex(handle)
            kernel32.CloseHandle(handle)
    except Exception:
        pass

def signal_existing_instance_to_show():
    kernel32.OpenEventW.restype = wintypes.HANDLE
    kernel32.OpenEventW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
    EVENT_MODIFY_STATE = 0x0002
    h = kernel32.OpenEventW(EVENT_MODIFY_STATE, False, SHOW_EVENT_NAME)
    if h:
        kernel32.SetEvent(h)
        kernel32.CloseHandle(h)
        return True
    return False

def create_auto_reset_event(name):
    kernel32.CreateEventW.restype = wintypes.HANDLE
    kernel32.CreateEventW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.BOOL, wintypes.LPCWSTR]
    return kernel32.CreateEventW(None, False, False, name)

def wait_for_event_and_call(hEvent, callback):
    def _runner():
        while True:
            res = kernel32.WaitForSingleObject(hEvent, 0xFFFFFFFF)  # INFINITE
            if res == 0:  # WAIT_OBJECT_0
                try:
                    callback()
                except Exception:
                    pass
            else:
                break
    t = threading.Thread(target=_runner, daemon=True)
    t.start()
    return t


class AudioMixerApp(ctk.CTk):
    FONT_TITLE = "Amazing Views"  # Nombre de familia interno del TTF
    FONT_UI = "Segoe UI"

    # ===== Precisión/fluidez =====
    ANIM_FPS = 60
    ANIM_EASING = 0.38

    BASE_PROCESS_INTERVAL_MS = 30
    VOLUME_TICK_MS = 20

    VOLUME_MIN_DELTA = 0.2
    VOLUME_MIN_INTERVAL_MS = 40

    SLIDER_MIN_DELTA = 0.1
    KNOB_MIN_DELTA = 0.2

    # Knobs inferiores
    KNOB_CANVAS_SIZE = 120
    KNOB_MARGIN = 10
    KNOB_STROKE = 4

    # Sliders superiores
    SLIDER_TRACK_WIDTH = 15
    SLIDER_HANDLE_WIDTH = 120
    SLIDER_LENGTH = 400
    SLIDER_BUTTON_LEN = 18
    SLIDER_BUTTON_RADIUS = 3
    SLIDER_HANDLE_CHAMFER = 150
    SLIDER_TOP_OFFSET = 0
    SLIDER_BOTTOM_OFFSET = 10

    ANALOG_MIN_RAW = 0      # Usar todo el rango del ADC evita zonas muertas en los extremos
    ANALOG_MAX_RAW = 1023
    ANALOG_CURVE_EXP = 0.8
    ANALOG_CALIBRATION_MIN_RANGE = 60
    ANALOG_INITIAL_MARGIN = 120  # margen para pre-calibrar el primer valor y evitar saltos iniciales
    ANALOG_BOOT_CENTER = 512     # valor central que asumimos hasta leer datos reales

    ZERO_EPS = 0.01
    ENDPOINT_CACHE_TTL_S = 10

    def __init__(self, mutex_handle=None, show_event_handle=None):
        super().__init__()
        self._mutex_handle = mutex_handle
        self._show_event_handle = show_event_handle
        self._show_event_thread = None

        try:
            comtypes.CoInitialize()
        except Exception:
            pass

        set_windows_app_user_model_id(APP_USER_MODEL_ID)

        # --- Cargar fuente privada antes de crear widgets que la usen ---
        self._font_path = resource_path("assets/AmazingViews.ttf")
        if os.path.exists(self._font_path):
            ok = load_private_font(self._font_path)
            if not ok:
                print("⚠ No se pudo registrar la fuente en memoria:", self._font_path)
        else:
            print("⚠ Fuente no encontrada (usará fallback):", self._font_path)

        # --- Iconos (ico + fallback PNG opcional) ---
        try:
            ico = resource_path("assets/knob.ico")
            if os.path.exists(ico):
                self.iconbitmap(ico)
        except Exception:
            pass
        try:
            png_fallback = resource_path("assets/knob_256.png")
            if os.path.exists(png_fallback):
                self.iconphoto(True, tk.PhotoImage(file=png_fallback))
        except Exception:
            pass

        self._anim_running = False
        self._anim_after_id = None

        self.themes = {
            "Negro": {"background": "#1c1c1c","panel":"#2E2E2E","panel_border":"#1F1F1F",
                      "fader_track":"#1A1A1A","knob_bg":"#1A1A1A","knob_outline":"#404040",
                      "control":"#9E9E9E","text":"#BDBDBD","hover":"#4A6572","handle":"#FFFFFF","accent":"#FFB347"},
            "Azul": {"background":"#0c0f2a","panel":"#1919e6","panel_border":"#0f1180",
                     "fader_track":"#1416b0","knob_bg":"#0f1180","knob_outline":"#6f73ff",
                     "control":"#c4c8ff","text":"#f7f8ff","hover":"#2f31ff","handle":"#000000","accent":"#ffb86c"},
            "Rojo": {"background":"#731111","panel":"#c41f1f","panel_border":"#4f0a0a",
                     "fader_track":"#8f1616","knob_bg":"#8f1616","knob_outline":"#f35a5a",
                     "control":"#ffb2b2","text":"#ffecec","hover":"#d93434","handle":"#000000","accent":"#4fc3f7"},
            "Blanco": {"background":"#E0E0E0","panel":"#F5F5F5","panel_border":"#BDBDBD",
                       "fader_track":"#E0E0E0","knob_bg":"#E0E0E0","knob_outline":"#BDBDBD",
                       "control":"#757575","text":"#212121","hover":"#BDBDBD","handle":"#000000","accent":"#3F51B5"},
            "Gris Claro": {"background":"#555555","panel":"#7a7a7a","panel_border":"#3d3d3d",
                           "fader_track":"#5f5f5f","knob_bg":"#5f5f5f","knob_outline":"#a5a5a5",
                           "control":"#d6d6d6","text":"#1e1e1e","hover":"#8a8a8a","handle":"#000000","accent":"#ffb300"},
            "Morado": {"background":"#1E0A33","panel":"#4d139e","panel_border":"#300a63",
                        "fader_track":"#1E0A33","knob_bg":"#1E0A33","knob_outline":"#7C3CBD",
                        "control":"#CDAAF5","text":"#F5ECFF","hover":"#7C3CBD","handle":"#FFFFFF","accent":"#FF9E7A"},
        }
        self.current_theme_name = "Negro"

        self.config_file = os.path.join(os.path.expanduser("~"), ".knob_config.json")
        self.error_log = os.path.join(os.path.expanduser("~"), "knob_error.log")

        ctk.set_appearance_mode("Dark")
        self.title(APP_TITLE)
        self.geometry("500x800")
        self.resizable(False, False)

        # ====== ventana keepalive (siempre escondida) ======
        self._keepalive = tk.Toplevel(self)
        self._keepalive.withdraw()
        self._keepalive.protocol("WM_DELETE_WINDOW", lambda: None)

        # Estado
        self.arduino = None
        self.listener_thread = None
        self.app_running = True
        self.assigned_processes = [None]*5
        self.data_queue = queue.Queue()
        self.ui_task_queue = queue.Queue()
        self.last_processed_volumes = [-1.0]*5

        self.sliders = []
        self.knobs = []
        self.app_name_labels = [None]*5
        self.knob_canvases = []
        self.selector_window = None
        self.settings_window = None
        self._settings_icon_photo = {}
        self.active_channel_frame = None
        self._raw_min = [float('inf')]*5
        self._raw_max = [float('-inf')]*5
        self._slider_handle_image_path = resource_path("assets/slider_handle.png")
        if not os.path.exists(self._slider_handle_image_path):
            self._slider_handle_image_path = None

        # Bandeja
        self.minimize_to_tray_var = ctk.BooleanVar(value=True)
        self.startup_with_windows_var = ctk.BooleanVar(value=self.check_startup_status())
        self.tray_icon = None
        self._tray_running = False
        self._restoring = False

        # Valores animados / envío
        boot_percent = self._boot_percent()
        self.current_values = [boot_percent]*5
        self.target_values = [boot_percent]*5
        self.last_sent_values = [boot_percent]*5
        self.last_sent_times = [0.0]*5
        self._first_value_applied = [False]*5  # evita animaciones bruscas al inicio

        self._last_serial_data_ts = time.time()
        self._endpoint_cache = None
        self._endpoint_cache_ts = 0.0
        self._mic_endpoint_cache = None
        self._mic_endpoint_cache_ts = 0.0

        # ---- UI ----
        self.status_label = ctk.CTkLabel(self, text="Iniciando...", font=ctk.CTkFont(size=14, family=self.FONT_UI))
        self.status_label.pack(pady=(0, 5))

        self.device_frame = ctk.CTkFrame(self, corner_radius=10, border_width=2)
        self.device_frame.pack(pady=10, padx=20, fill="both", expand=True)

        panel_color = self.themes[self.current_theme_name]["panel"]
        initial_text_color = "#3A3A3A" if self.current_theme_name == "Blanco" else "#FFFFFF"
        self.settings_button = ctk.CTkButton(self.device_frame, text="⚙", width=30, height=30,
                     font=ctk.CTkFont(size=20), fg_color=panel_color,
                     hover_color=panel_color, text_color=initial_text_color,
                     command=self.open_settings_window)
        self.settings_button.place(relx=1.0, x=-15, y=15, anchor="ne")

        self.knob_title = ctk.CTkLabel(self.device_frame, text="KNOB",
                                       font=ctk.CTkFont(size=60, family=self.FONT_TITLE, weight="bold"))
        self.knob_title.pack(pady=(15, 5))

        # Sliders arriba (equidistantes)
        sliders_container = ctk.CTkFrame(self.device_frame, fg_color="transparent")
        sliders_container.pack(pady=(5, 10), padx=30, fill="both", expand=True)
        sliders_container.grid_columnconfigure((0, 1, 2), weight=1, uniform="sliders")
        sliders_container.grid_rowconfigure(0, weight=1)
        for i in range(3):
            self.create_slider_channel(sliders_container, i)

        # Knobs abajo (dos grandes)
        knobs_container = ctk.CTkFrame(self.device_frame, fg_color="transparent")
        knobs_container.pack(pady=(10, 20), padx=30, fill="x", side="bottom")
        knobs_container.grid_columnconfigure((0, 1), weight=1, uniform="knobs")
        for i in range(2):
            self.create_knob_channel(knobs_container, i)

        self.load_settings()
        self.after(0, self.post_init)
        self.after(40, self._process_ui_tasks)

        # La X SIEMPRE manda a bandeja
        self.protocol("WM_DELETE_WINDOW", self.on_closing_to_tray)

        if self._show_event_handle:
            self._show_event_thread = wait_for_event_and_call(self._show_event_handle, self.restore_from_tray_async)

    def invoke_ui(self, callback, *args, **kwargs):
        if not callable(callback):
            return
        self.ui_task_queue.put((callback, args, kwargs))

    def _process_ui_tasks(self):
        try:
            while True:
                callback, args, kwargs = self.ui_task_queue.get_nowait()
                try:
                    callback(*args, **kwargs)
                except Exception as task_error:
                    self._log_error(f"_process_ui_tasks: {task_error}")
        except queue.Empty:
            pass
        finally:
            if self.app_running:
                try:
                    self.after(40, self._process_ui_tasks)
                except Exception:
                    pass

    def update_status_async(self, text=None, text_color=None, font=None):
        def _apply():
            kwargs = {}
            if text is not None:
                kwargs["text"] = text
            if text_color is not None:
                kwargs["text_color"] = text_color
            if font is not None:
                kwargs["font"] = font
            if kwargs:
                self.status_label.configure(**kwargs)
        self.invoke_ui(_apply)

    def _get_settings_icon_photo(self, size=48, color="#3A3A3A"):
        if Image is None or ImageTk is None:
            return None
        key = (size, color)
        if key not in self._settings_icon_photo:
            self._settings_icon_photo[key] = self._create_gear_photo(size=size, color=color)
        return self._settings_icon_photo[key]

    def _create_gear_photo(self, size=48, color="#3A3A3A"):
        if Image is None or ImageTk is None:
            return None
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        center = size / 2.0
        outer_radius = size * 0.32
        tooth_outer = size * 0.48
        tooth_width = max(2, int(size * 0.12))
        inner_radius = size * 0.18

        for angle_deg in range(0, 360, 45):
            angle = math.radians(angle_deg)
            x_inner = center + math.cos(angle) * outer_radius
            y_inner = center + math.sin(angle) * outer_radius
            x_outer = center + math.cos(angle) * tooth_outer
            y_outer = center + math.sin(angle) * tooth_outer
            draw.line((x_inner, y_inner, x_outer, y_outer), fill=color, width=tooth_width)

        draw.ellipse((center - outer_radius,
                      center - outer_radius,
                      center + outer_radius,
                      center + outer_radius), fill=color)
        draw.ellipse((center - inner_radius,
                      center - inner_radius,
                      center + inner_radius,
                      center + inner_radius), fill=(0, 0, 0, 0))

        return ImageTk.PhotoImage(img)

    # ======= UI canales =======
    def _format_display_name(self, name: str) -> str:
        if not name:
            return "No Asignado"
        base = name[:-4] if name.lower().endswith(".exe") else name
        return base[:1].upper() + base[1:].lower()

    def create_slider_channel(self, parent, channel_index):
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.grid(row=0, column=channel_index, padx=20, sticky="ns")

        name_frame = ctk.CTkFrame(col, corner_radius=6, cursor="hand2")
        name_frame.pack(pady=5, padx=5, fill="x")

        app_name_label = ctk.CTkLabel(name_frame, text="No Asignado",
                                      font=ctk.CTkFont(size=12, family=self.FONT_UI),
                                      wraplength=120, fg_color="transparent")
        app_name_label.pack(pady=4, padx=8, fill="x")
        self.app_name_labels[channel_index] = app_name_label

        name_frame.bind("<Enter>", lambda e, f=name_frame: self.on_name_frame_enter(f))
        name_frame.bind("<Leave>", lambda e, f=name_frame: self.on_name_frame_leave(f))
        name_frame.bind("<Button-1>", lambda e, idx=channel_index: self.open_process_selector(idx))
        app_name_label.bind("<Button-1>", lambda e, idx=channel_index: self.open_process_selector(idx))

        fader_container = ctk.CTkFrame(col, fg_color="transparent")
        fader_container.pack(pady=10, fill="both", expand=True)
        fader_container.pack_propagate(False)
        fader_container.bind("<Configure>", lambda e, slider_index=channel_index: self._position_slider(slider_index))

        fader = WideHandleSlider(
            fader_container,
            from_=0,
            to=100,
            orientation="vertical",
            state="disabled",
            width=self.SLIDER_HANDLE_WIDTH,
            height=self.SLIDER_LENGTH + self.SLIDER_BUTTON_LEN,
            corner_radius=0,
            button_length=self.SLIDER_BUTTON_LEN,
            button_corner_radius=self.SLIDER_BUTTON_RADIUS,
            track_width=self.SLIDER_TRACK_WIDTH,
            track_end_margin=0,
            handle_chamfer=self.SLIDER_HANDLE_CHAMFER,
            handle_image_path=self._slider_handle_image_path,
        )
        fader.set(0)
        self.sliders.append(fader)
        self._position_slider(channel_index)

    def create_knob_channel(self, parent, knob_index):
        idx = knob_index + 3
        col = ctk.CTkFrame(parent, fg_color="transparent")
        col.grid(row=0, column=knob_index, padx=16, sticky="ns")

        name_frame = ctk.CTkFrame(col, corner_radius=6, cursor="hand2")
        name_frame.pack(pady=6, padx=6, fill="x")

        app_name_label = ctk.CTkLabel(name_frame, text="No Asignado",
                                      font=ctk.CTkFont(size=12, family=self.FONT_UI),
                                      wraplength=160, fg_color="transparent")
        app_name_label.pack(pady=5, padx=10, fill="x")
        self.app_name_labels[idx] = app_name_label

        name_frame.bind("<Enter>", lambda e, f=name_frame: self.on_name_frame_enter(f))
        name_frame.bind("<Leave>", lambda e, f=name_frame: self.on_name_frame_leave(f))
        name_frame.bind("<Button-1>", lambda e, i=idx: self.open_process_selector(i))
        app_name_label.bind("<Button-1>", lambda e, i=idx: self.open_process_selector(i))

        size = self.KNOB_CANVAS_SIZE
        margin = self.KNOB_MARGIN
        stroke = self.KNOB_STROKE

        knob_canvas = ctk.CTkCanvas(col, width=size, height=size, highlightthickness=0)
        knob_canvas.pack(pady=12)
        self.knob_canvases.append(knob_canvas)

        oval_id = knob_canvas.create_oval(margin, margin, size - margin, size - margin, width=stroke)
        cx = cy = size / 2.0
        marker_id = knob_canvas.create_line(cx, cy, cx, margin + stroke, width=stroke, capstyle="round")

        self.knobs.append({'canvas': knob_canvas, 'marker_id': marker_id, 'oval_id': oval_id,
                           'size': size, 'margin': margin, 'stroke': stroke})

    def _position_slider(self, slider_index):
        try:
            slider = self.sliders[slider_index]
        except IndexError:
            return

        container = slider.master
        if container is None:
            return

        container.update_idletasks()
        container_height = container.winfo_height()
        if container_height <= 0:
            self.after(10, lambda idx=slider_index: self._position_slider(idx))
            return

        available = max(0, container_height - (self.SLIDER_TOP_OFFSET + self.SLIDER_BOTTOM_OFFSET))
        if available <= 0:
            target_height = 1
        else:
            gap = int(self.SLIDER_BUTTON_LEN * 0.25)
            target_height = max(1, available - gap)  # recorta un poco para evitar que se pase abajo
        y_offset = self.SLIDER_TOP_OFFSET

        if not slider.place_info():
            slider.place(relx=0.5, rely=0.0, anchor="n")

        slider.place_configure(y=y_offset, height=target_height)

    # ======= Hover =======
    def on_name_frame_enter(self, frame):
        theme = self.themes[self.current_theme_name]
        frame.configure(fg_color=theme["hover"])
        for w in frame.winfo_children():
            if isinstance(w, ctk.CTkLabel):
                w.configure(fg_color=theme["hover"])

    def on_name_frame_leave(self, frame):
        if frame == self.active_channel_frame:
            return
        theme = self.themes[self.current_theme_name]
        track_color = theme["fader_track"]
        frame.configure(fg_color=track_color)
        for w in frame.winfo_children():
            if isinstance(w, ctk.CTkLabel):
                w.configure(fg_color=track_color)

    # ======= Tema / Settings =======
    def apply_theme(self, theme_name):
        theme = self.themes.get(theme_name)
        if not theme:
            return
        self.current_theme_name = theme_name

        self.configure(fg_color=theme["background"])
        self.status_label.configure(text_color=theme["text"])
        self.device_frame.configure(fg_color=theme["panel"], border_color=theme["panel_border"])
        self.knob_title.configure(text_color=theme["text"])
        button_text_color = "#3A3A3A" if theme_name == "Blanco" else "#FFFFFF"
        self.settings_button.configure(fg_color=theme["panel"], hover_color=theme["panel"],
                   text_color=button_text_color)

        for label in self.app_name_labels:
            if label:
                frame = label.master
                color = theme["hover"] if frame == self.active_channel_frame else theme["fader_track"]
                frame.configure(fg_color=color)
                label.configure(fg_color=color, text_color=theme["text"])

        handle_color = theme.get("handle", "#FFFFFF")
        track_color = "#000000"
        for slider in self.sliders:
            slider.configure(button_color=handle_color, progress_color=track_color,
                             fg_color=track_color, button_hover_color=theme["hover"])

        for knob in self.knobs:
            canvas = knob['canvas']
            canvas.configure(bg=theme["panel"])
            canvas.itemconfig(knob['oval_id'], fill=theme["knob_bg"], outline=theme["knob_outline"],
                              width=knob.get('stroke', self.KNOB_STROKE))
            canvas.itemconfig(knob['marker_id'], fill=handle_color, width=knob.get('stroke', self.KNOB_STROKE))

        self.save_settings()

    def open_settings_window(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.focus()
            return

        self.settings_window = ctk.CTkToplevel(self)
        icon_color = "#3A3A3A" if self.current_theme_name == "Blanco" else "#FFFFFF"
        gear_photo = self._get_settings_icon_photo(color=icon_color)
        if gear_photo:
            try:
                self.settings_window.iconphoto(False, gear_photo)
            except Exception:
                pass
        self.settings_window.title("Configuracion")
        self.settings_window.geometry("400x350")
        self.settings_window.transient(self)
        self.settings_window.grab_set()

        settings_frame = ctk.CTkFrame(self.settings_window)
        settings_frame.pack(expand=True, fill="both", padx=10, pady=10)

        self.startup_with_windows_var.set(self.check_startup_status())

        startup_checkbox = ctk.CTkCheckBox(settings_frame, text="Iniciar con Windows",
                                           font=ctk.CTkFont(family=self.FONT_UI, size=12),
                                           variable=self.startup_with_windows_var,
                                           onvalue=True, offvalue=False,
                                           command=self.toggle_startup)
        startup_checkbox.pack(pady=10, padx=10, anchor="w")
        if winshell is None or Dispatch is None:
            self.startup_with_windows_var.set(False)
            startup_checkbox.configure(state="disabled")
            ctk.CTkLabel(settings_frame, text="Instala 'winshell' y 'pywin32' para esta opcion", text_color="orange").pack(
                pady=5, padx=10, anchor="w")

        minimize_checkbox = ctk.CTkCheckBox(settings_frame, text="Minimizar a la bandeja al cerrar",
                                            variable=self.minimize_to_tray_var, onvalue=True, offvalue=False,
                                            font=ctk.CTkFont(family=self.FONT_UI, size=12),
                                            command=self.save_settings)
        minimize_checkbox.pack(pady=10, padx=10, anchor="w")
        if pystray is None:
            minimize_checkbox.configure(state="disabled")
            ctk.CTkLabel(settings_frame, text="Instala 'pystray' y 'Pillow' para esta opcion", text_color="orange").pack(
                pady=5, padx=10, anchor="w")

        ctk.CTkFrame(settings_frame, height=2).pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(settings_frame, text="Tema de Color:", font=ctk.CTkFont(family=self.FONT_UI, size=12)).pack(
            pady=(10, 5), padx=10, anchor="w")

        color_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        color_frame.pack(fill="x", padx=10)
        for i, (name, theme) in enumerate(self.themes.items()):
            ctk.CTkButton(color_frame, text="", fg_color=theme["panel"], width=40, height=30, corner_radius=8,
                          border_width=2, border_color=theme["panel_border"],
                          command=lambda n=name: self.apply_theme(n)).grid(row=0, column=i, padx=5)

    def toggle_startup(self):
        desired_state = bool(self.startup_with_windows_var.get())

        if winshell is None or Dispatch is None:
            self.startup_with_windows_var.set(False)
            messagebox.showerror("Iniciar con Windows", "Instala los paquetes 'winshell' y 'pywin32' para usar esta opcion.")
            return

        shortcut_path = self.get_shortcut_path()
        shortcut_dir = os.path.dirname(shortcut_path)
        error_message = None

        try:
            if desired_state:
                if shortcut_dir and not os.path.exists(shortcut_dir):
                    os.makedirs(shortcut_dir, exist_ok=True)

                target_path = sys.executable
                script_path = os.path.abspath(__file__)
                working_dir = os.path.dirname(script_path)

                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target_path
                if not getattr(sys, 'frozen', False):
                    shortcut.Arguments = f'"{script_path}"'
                else:
                    shortcut.Arguments = ""
                shortcut.WorkingDirectory = working_dir

                ico = resource_path("assets/knob.ico")
                if os.path.exists(ico):
                    shortcut.IconLocation = ico
                shortcut.save()
            else:
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
        except Exception as e:
            error_message = str(e)
            self._log_error(f"toggle_startup: {e}")

        if error_message:
            self.startup_with_windows_var.set(not desired_state)
            messagebox.showerror("Iniciar con Windows",
                                 "No se pudo actualizar el inicio automatico. Revisa los permisos e intenta de nuevo.")

    def check_startup_status(self):
        if winshell is None:
            return False
        return os.path.exists(self.get_shortcut_path())

    def get_shortcut_path(self):
        if winshell is None:
            return os.path.join(os.path.expanduser("~"), "Start Menu", "Programs", "Startup", "KNOB.lnk")
        return os.path.join(winshell.startup(), "KNOB.lnk")

    def save_settings(self):
        settings = {"theme": self.current_theme_name,
                    "minimize_to_tray": self.minimize_to_tray_var.get(),
                    "assigned_processes": self.assigned_processes}
        try:
            with open(self.config_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            self._log_error(f"Error al guardar configuracion: {e}")

    def load_settings(self):
        try:
            with open(self.config_file, 'r') as f:
                settings = json.load(f)
            self.minimize_to_tray_var.set(settings.get("minimize_to_tray", True))
            loaded_processes = settings.get("assigned_processes", [None]*5)
            for i, process_name in enumerate(loaded_processes):
                if process_name:
                    self.assign_process_to_channel(process_name, i)
            self.apply_theme(settings.get("theme", "Negro"))
        except (FileNotFoundError, json.JSONDecodeError):
            self.apply_theme("Negro")
        except Exception as e:
            self._log_error(f"Error al cargar configuracion: {e}")

    # ======= Selector =======
    def open_process_selector(self, channel_index):
        if self.selector_window is not None and self.selector_window.winfo_exists():
            self.selector_window.focus(); return

        if self.active_channel_frame:
            self.on_name_frame_leave(self.active_channel_frame)

        current_frame = self.app_name_labels[channel_index].master
        self.on_name_frame_enter(current_frame)
        self.active_channel_frame = current_frame

        self.selector_window = ctk.CTkToplevel(self)
        self.selector_window.title("Seleccionar Aplicacion")
        self.selector_window.geometry("400x500")
        self.selector_window.transient(self)
        self.selector_window.grab_set()
        self.selector_window.protocol("WM_DELETE_WINDOW", self.on_selector_close)

        scrollable = ctk.CTkScrollableFrame(self.selector_window, label_text="Asignar Canal",
                                            label_font=ctk.CTkFont(family=self.FONT_UI, size=14))
        scrollable.pack(expand=True, fill="both", padx=10, pady=10)

        special_options = ["Volumen General", "Aplicacion en Foco", "Microfono", "Quitar Asignacion"]
        for option in special_options:
            ctk.CTkButton(scrollable, text=option,
                          font=ctk.CTkFont(family=self.FONT_UI, size=12, weight="bold"),
                          command=lambda p=option, idx=channel_index: self.assign_process_from_selector(p, idx)
                          ).pack(pady=4, padx=10, fill="x")

        ctk.CTkFrame(scrollable, height=2).pack(fill="x", padx=10, pady=10)

        plist = self.get_audio_processes()
        for pname in plist:
            if pname == "No hay procesos":
                continue
            nice = self._format_display_name(pname)
            ctk.CTkButton(scrollable, text=nice, font=ctk.CTkFont(family=self.FONT_UI, size=12),
                          command=lambda p=pname, idx=channel_index: self.assign_process_from_selector(p, idx)
                          ).pack(pady=4, padx=10, fill="x")

    def on_selector_close(self):
        if self.active_channel_frame:
            theme = self.themes[self.current_theme_name]
            track = theme["fader_track"]
            self.active_channel_frame.configure(fg_color=track)
            for w in self.active_channel_frame.winfo_children():
                if isinstance(w, ctk.CTkLabel): w.configure(fg_color=track)
            self.active_channel_frame = None
        if self.selector_window:
            self.selector_window.destroy()
            self.selector_window = None

    def assign_process_from_selector(self, process_name, channel_index):
        if process_name == "Quitar Asignacion":
            self.assign_process_to_channel(None, channel_index)
        else:
            self.assign_process_to_channel(process_name, channel_index)
        self.on_selector_close()

    def assign_process_to_channel(self, process_name, channel_index):
        self.assigned_processes[channel_index] = process_name
        display_name = self._format_display_name(process_name) if process_name else "No Asignado"
        self.app_name_labels[channel_index].configure(text=display_name)
        self.save_settings()

    def get_audio_processes(self):
        names = set()
        try:
            for s in AudioUtilities.GetAllSessions():
                if s.Process and s.Process.name().lower() not in ["svchost.exe", "audiodg.exe"]:
                    names.add(s.Process.name())
        except Exception:
            pass
        return sorted(list(names)) if names else ["No hay procesos"]

    # ======= Arduino =======
    def init_arduino_thread(self):
        self.listener_thread = threading.Thread(target=self.arduino_connection_manager, daemon=True)
        self.listener_thread.start()

    def _is_candidate_port(self, pinfo):
        try:
            if pinfo.vid and pinfo.pid and (pinfo.vid, pinfo.pid) in KNOWN_DEVICES:
                return True
        except Exception:
            pass
        desc = (pinfo.description or "").lower()
        hwid = (pinfo.hwid or "").lower()
        if any(k in desc for k in ["arduino","ch340","wch","usb-serial","cp210"]): return True
        if any(k in hwid for k in ["arduino","ch340","wch","cp210"]): return True
        return False

    def _try_open_and_handshake(self, port_device):
        try:
            ser = serial.Serial(port_device, BAUD_RATE, timeout=SERIAL_READ_TIMEOUT)
        except Exception as e:
            self._log_error(f"_try_open: no se pudo abrir {port_device}: {e}")
            return None
        try:
            try:
                ser.setDTR(False); ser.setRTS(False); time.sleep(0.05)
                ser.setDTR(True); ser.setRTS(True)
            except Exception: pass
            try:
                ser.reset_input_buffer(); ser.reset_output_buffer()
            except Exception: pass

            start_ts = time.time()
            saw_await = False; saw_ok = False; saw_data = False

            while time.time() - start_ts < HANDSHAKE_WINDOW_S:
                line = ser.readline().decode('utf-8', errors='ignore').strip()
                if not line:
                    try: ser.write(INIT_TOKEN_START.encode("utf-8"))
                    except Exception: pass
                    continue
                if INIT_TOKEN_AWAIT in line:
                    saw_await = True
                    try: ser.write(INIT_TOKEN_START.encode("utf-8"))
                    except Exception: pass
                    continue
                if INIT_TOKEN_OK in line:
                    saw_ok = True; break
                if FRAME_SEP in line:
                    parts = line.split(FRAME_SEP)
                    if len(parts) == FRAME_VALUES:
                        try:
                            _ = [int(x) for x in parts]
                            saw_data = True; break
                        except Exception:
                            pass

            if saw_ok or saw_data or (saw_await and not saw_ok):
                if saw_await and not saw_ok:
                    try: ser.write(INIT_TOKEN_START.encode("utf-8"))
                    except Exception: pass
                try: ser.reset_input_buffer()
                except Exception: pass
                return ser

        except Exception as e:
            self._log_error(f"_try_open_and_handshake error {port_device}: {e}")

        try: ser.close()
        except Exception: pass
        return None

    def arduino_connection_manager(self):
        backoff_s = 2
        while self.app_running:
            if not self.arduino or not self.arduino.is_open:
                self.update_status_async(text="Buscando Arduino...", text_color=None)
                try: ports = serial.tools.list_ports.comports()
                except Exception as e:
                    self._log_error(f"Error listando puertos: {e}"); ports = []

                candidates = [p for p in ports if self._is_candidate_port(p)] or list(ports)
                for p in candidates:
                    self.update_status_async(text=f"Probando puerto {p.device}...", text_color=None)
                    ser = self._try_open_and_handshake(p.device)
                    if ser:
                        self.arduino = ser
                        self._last_serial_data_ts = time.time()
                        self.update_status_async(text=f"Conectado a Arduino en {self.arduino.port}", text_color="light green")
                        self.invoke_ui(lambda: self.after(3000, self.show_connection_indicator))
                        break

                if not self.arduino:
                    self.update_status_async(text="No se pudo encontrar el Arduino. Reintentando...", text_color=None)
                    time.sleep(backoff_s); backoff_s = min(backoff_s * 2, 10)
                    continue
                else:
                    backoff_s = 2

            try:
                self.listen_for_data()
            except Exception as e:
                self._log_error(f"listen_for_data (outer): {e}")

            try:
                if self.arduino: self.arduino.close()
            except Exception: pass
            self.arduino = None
            time.sleep(1)

    def show_connection_indicator(self):
        if self.arduino and self.arduino.is_open:
            self.status_label.configure(text="✔ Conectado",
                                        font=ctk.CTkFont(size=16, family=self.FONT_UI),
                                        text_color="light green")

    def listen_for_data(self):
        while self.app_running and self.arduino and self.arduino.is_open:
            idle = time.time() - self._last_serial_data_ts
            if idle >= NO_DATA_RECONNECT_S:
                self.update_status_async(text="Sin datos del dispositivo. Reintentando...", text_color="orange")
                break
            try:
                if self.arduino.in_waiting > 0:
                    data_line = self.arduino.readline().decode('utf-8', errors='ignore').strip()
                    if data_line: self._last_serial_data_ts = time.time()
                    if FRAME_SEP in data_line:
                        self.data_queue.put(data_line)
                else:
                    time.sleep(0.01)
            except serial.SerialException:
                self.update_status_async(text="Conexion perdida.", text_color="orange"); break
            except Exception as e:
                self._log_error(f"listen_for_data: {e}"); break

    # ======= Intervalos dinámicos =======
    @property
    def PROCESS_INTERVAL_MS(self):
        base = type(self).BASE_PROCESS_INTERVAL_MS
        try:
            if not self.winfo_viewable(): return max(120, base)
        except Exception:
            pass
        return base

    @property
    def ANIM_INTERVAL_MS(self):
        base = int(1000 / self.ANIM_FPS)
        try:
            if not self.winfo_viewable(): return max(50, int(base * 1.5))
        except Exception:
            pass
        return base

    # ======= Bucle principal =======
    def post_init(self):
        try: self.init_arduino_thread()
        except Exception as e: self._log_error(f"post_init init_arduino_thread: {e}")
        try: self.schedule_process_serial_queue()
        except Exception as e: self._log_error(f"post_init schedule_process_serial_queue: {e}")
        try: self.schedule_volume_tick()
        except Exception as e: self._log_error(f"post_init schedule_volume_tick: {e}")

    def schedule_process_serial_queue(self):
        try: self.process_serial_queue()
        except Exception as e: self._log_error(f"process_serial_queue (run): {e}")
        finally:
            try: self.after(self.PROCESS_INTERVAL_MS, self.schedule_process_serial_queue)
            except Exception as e: self._log_error(f"process_serial_queue (after): {e}")

    def process_serial_queue(self):
        changed = False
        while not self.data_queue.empty():
            data_line = self.data_queue.get_nowait()
            values = data_line.split(FRAME_SEP)
            if len(values) == FRAME_VALUES:
                for i, val_str in enumerate(values):
                    try:
                        # --- Mapeo más preciso ---
                        percent = self._analog_to_percent(i, int(val_str))
                        percent = max(0.0, min(100.0, percent))

                        if not self._first_value_applied[i]:
                            self._first_value_applied[i] = True
                            self.last_processed_volumes[i] = percent
                            self.current_values[i] = percent
                            self.target_values[i] = percent
                            if i < 3:
                                if i < len(self.sliders):
                                    self.sliders[i].set(percent)
                            else:
                                self.update_knob_visual(i - 3, percent)
                            continue

                        if abs(percent - self.last_processed_volumes[i]) > 0.5:
                            self.last_processed_volumes[i] = percent
                            self.target_values[i] = percent
                            changed = True
                    except (ValueError, IndexError):
                        continue
        if changed:
            self.ensure_animating()

    # ======= Curva de calibración =======
    def _analog_to_percent(self, channel_index, raw_value):
        """Convierte la lectura cruda al rango completo 0–100 % con auto-calibración."""
        raw = max(self.ANALOG_MIN_RAW, min(self.ANALOG_MAX_RAW, float(raw_value)))

        normalized = None
        if 0 <= channel_index < len(self._raw_min):
            if math.isinf(self._raw_min[channel_index]):
                base_min = self.ANALOG_MIN_RAW
                base_max = self.ANALOG_MAX_RAW
                self._raw_min[channel_index] = base_min
                self._raw_max[channel_index] = base_max

            self._raw_min[channel_index] = min(self._raw_min[channel_index], raw)
            self._raw_max[channel_index] = max(self._raw_max[channel_index], raw)
            span = self._raw_max[channel_index] - self._raw_min[channel_index]
            if span >= self.ANALOG_CALIBRATION_MIN_RANGE:
                normalized = (raw - self._raw_min[channel_index]) / max(1.0, span)

        if normalized is None:
            span = max(1.0, float(self.ANALOG_MAX_RAW - self.ANALOG_MIN_RAW))
            normalized = (raw - self.ANALOG_MIN_RAW) / span

        normalized = max(0.0, min(1.0, normalized))
        corrected = math.pow(normalized, self.ANALOG_CURVE_EXP)
        return corrected * 100.0

    def _boot_percent(self):
        span = max(1.0, float(self.ANALOG_MAX_RAW - self.ANALOG_MIN_RAW))
        normalized = (self.ANALOG_BOOT_CENTER - self.ANALOG_MIN_RAW) / span
        normalized = max(0.0, min(1.0, normalized))
        return math.pow(normalized, self.ANALOG_CURVE_EXP) * 100.0

    # ======= Animación =======
    def ensure_animating(self):
        if not self._anim_running:
            self._anim_running = True
            self._schedule_next_anim_frame()

    def _schedule_next_anim_frame(self):
        if self._anim_after_id:
            try: self.after_cancel(self._anim_after_id)
            except Exception: pass
        self._anim_after_id = self.after(self.ANIM_INTERVAL_MS, self.animate_controls)

    def animate_controls(self):
        try:
            any_active = False
            for i in range(5):
                current, target = self.current_values[i], self.target_values[i]
                if abs(current - target) < 0.05:
                    new_value = target
                else:
                    new_value = current + (target - current) * self.ANIM_EASING
                    any_active = True

                if i < 3:
                    if abs(new_value - self.current_values[i]) >= self.SLIDER_MIN_DELTA:
                        self.sliders[i].set(new_value)
                else:
                    self.update_knob_visual(i - 3, new_value)
                self.current_values[i] = new_value

            if any_active: self._schedule_next_anim_frame()
            else:
                self._anim_running = False
                self._anim_after_id = None
        except Exception as e:
            self._log_error(f"animate_controls: {e}")
            self._anim_running = False
            self._anim_after_id = None

    def update_knob_visual(self, knob_index, percent_value):
        knob_info = self.knobs[knob_index]
        canvas = knob_info['canvas']
        marker_id = knob_info['marker_id']
        size = knob_info.get('size', self.KNOB_CANVAS_SIZE)
        margin = knob_info.get('margin', self.KNOB_MARGIN)

        min_angle, max_angle = -135, 135
        angle = min_angle + (percent_value / 100.0) * (max_angle - min_angle)
        rad = math.radians(angle - 90)

        cx = cy = size / 2.0
        r = (size / 2.0) - margin - 2
        ex = cx + r * math.cos(rad)
        ey = cy + r * math.sin(rad)
        canvas.coords(marker_id, cx, cy, ex, ey)

    # ======= Envío de volumen =======
    def schedule_volume_tick(self):
        try: self.volume_tick()
        except Exception as e: self._log_error(f"volume_tick (run): {e}")
        finally:
            try: self.after(self.VOLUME_TICK_MS, self.schedule_volume_tick)
            except Exception as e: self._log_error(f"volume_tick (after): {e}")

    def _collect_audio_sessions(self):
        try:
            return AudioUtilities.GetAllSessions()
        except Exception as e:
            self._log_error(f"_collect_audio_sessions error: {e}")
            return []

    def _get_master_endpoint(self):
        now = time.time()
        if self._endpoint_cache and (now - self._endpoint_cache_ts) < self.ENDPOINT_CACHE_TTL_S:
            return self._endpoint_cache
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            endpoint = cast(interface, POINTER(IAudioEndpointVolume))
            self._endpoint_cache = endpoint
            self._endpoint_cache_ts = now
            return endpoint
        except Exception as e:
            self._endpoint_cache = None
            self._endpoint_cache_ts = 0.0
            self._log_error(f"_get_master_endpoint error: {e}")
            return None

    def _get_microphone_endpoint(self):
        now = time.time()
        if self._mic_endpoint_cache and (now - self._mic_endpoint_cache_ts) < self.ENDPOINT_CACHE_TTL_S:
            return self._mic_endpoint_cache
        try:
            devices = AudioUtilities.GetMicrophone()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            endpoint = cast(interface, POINTER(IAudioEndpointVolume))
            self._mic_endpoint_cache = endpoint
            self._mic_endpoint_cache_ts = now
            return endpoint
        except Exception as e:
            self._mic_endpoint_cache = None
            self._mic_endpoint_cache_ts = 0.0
            self._log_error(f"_get_microphone_endpoint error: {e}")
            return None

    def volume_tick(self):
        now = time.time()
        sessions_cache = None
        for i in range(5):
            proc = self.assigned_processes[i]
            if not proc:
                continue
            desired = self.current_values[i]
            last_val = self.last_sent_values[i]
            last_t = self.last_sent_times[i]

            if (abs(desired - last_val) >= self.VOLUME_MIN_DELTA) and ((now - last_t) * 1000 >= self.VOLUME_MIN_INTERVAL_MS):
                session_bundle = None
                if proc != "Volumen General":
                    if sessions_cache is None:
                        sessions_cache = self._collect_audio_sessions()
                    session_bundle = sessions_cache

                self._apply_volume(proc, desired, session_bundle)
                self.last_sent_values[i] = desired
                self.last_sent_times[i] = now

    def _apply_volume(self, process_name, volume_percent, sessions=None):
        """
        Control de volumen continuo y preciso con mute automático al 1 % o menos.
        """
        try:
            # Escala lineal
            scalar = max(0.0, min(1.0, volume_percent / 100.0))
            should_mute = scalar <= 0.01  # Mute si volumen ≤ 1 %

            # === 1) Volumen general ===
            if process_name == "Volumen General":
                endpoint = self._get_master_endpoint()
                if endpoint:
                    try:
                        endpoint.SetMute(should_mute, None)
                    except Exception:
                        pass
                    endpoint.SetMasterVolumeLevelScalar(0.0 if should_mute else scalar, None)
                return

            if process_name == "Microfono":
                endpoint = self._get_microphone_endpoint()
                if endpoint:
                    try:
                        endpoint.SetMute(should_mute, None)
                    except Exception:
                        pass
                    endpoint.SetMasterVolumeLevelScalar(0.0 if should_mute else scalar, None)
                return

            sessions = sessions or []

            # === 2) Aplicación en foco ===
            if process_name == "Aplicacion en Foco" and win32gui and win32process:
                hwnd = win32gui.GetForegroundWindow()
                if not hwnd:
                    return
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    if pid == os.getpid():
                        return
                except Exception:
                    pass

                exe_base = None
                try:
                    p = psutil.Process(pid)
                    exe_base = _norm_proc_name(p.name() or "")
                except Exception:
                    exe_base = None

                for session in sessions:
                    try:
                        if not session.Process:
                            continue
                        spid = session.Process.pid
                        sbase = _norm_proc_name(session.Process.name() or "")
                        if spid == pid or (exe_base and sbase == exe_base):
                            vol = session._ctl.QueryInterface(ISimpleAudioVolume)
                            vol.SetMute(should_mute, None)
                            vol.SetMasterVolume(0.0 if should_mute else scalar, None)
                    except Exception:
                        continue
                return

            # === 3) Aplicación específica ===
            target_base = _norm_proc_name(process_name)
            if not target_base:
                return

            for session in sessions:
                try:
                    if session.Process and _names_match(session.Process.name(), target_base):
                        vol = session._ctl.QueryInterface(ISimpleAudioVolume)
                        vol.SetMute(should_mute, None)
                        vol.SetMasterVolume(0.0 if should_mute else scalar, None)
                except Exception:
                    continue

        except Exception as e:
            self._log_error(f"_apply_volume mute≤1 error: {e}")

    # ======= Bandeja =======
    def on_closing_to_tray(self):
        if self.minimize_to_tray_var.get():
            self.hide_to_tray()
        else:
            self.shutdown_app()

    def _build_tray_icon(self):
        tray_image = None
        try:
            ico_path = resource_path("assets/knob.ico")
            if Image is not None and os.path.exists(ico_path):
                img = Image.open(ico_path)
                if img.mode not in ("RGB", "RGBA"):
                    img = img.convert("RGBA")
                size = min(img.width, img.height)
                img = ImageOps.fit(img, (size, size))
                tray_image = img
        except Exception:
            tray_image = None

        if tray_image is None and Image is not None:
            theme = self.themes[self.current_theme_name]
            tray_image = Image.new('RGB', (64, 64), theme["panel"])
            dc = ImageDraw.Draw(tray_image)
            dc.ellipse([(16, 16), (48, 48)], fill=theme["control"])

        if tray_image is None or pystray is None:
            return None

        menu = pystray.Menu(
            pystray.MenuItem('Mostrar', self.show_from_tray, default=True),
            pystray.MenuItem('Salir', self.quit_from_tray)
        )
        return pystray.Icon("AudioMixerApp", tray_image, "KNOB", menu)

    def hide_to_tray(self):
        try:
            self.withdraw()
        except Exception:
            self.iconify()

        if pystray is None or Image is None:
            return

        if self._tray_running and self.tray_icon:
            return

        self.tray_icon = self._build_tray_icon()
        if not self.tray_icon:
            return

        self._tray_running = True
        try:
            self.tray_icon.run_detached()
        except Exception:
            self._tray_running = False

    def restore_from_tray_async(self):
        if self._restoring:
            return
        self._restoring = True
        def _schedule_restore():
            try:
                self.after(0, self.restore_from_tray)
            except Exception:
                self._restoring = False
        self.invoke_ui(_schedule_restore)

    def restore_from_tray(self):
        try:
            self.deiconify()
            self.lift()
            self.focus_force()
        except Exception:
            try: self.state('normal')
            except Exception: pass

        if self.tray_icon:
            try: self.tray_icon.stop()
            except Exception: pass

        self.tray_icon = None
        self._tray_running = False
        self._restoring = False

    def show_from_tray(self, icon, item):
        try:
            self.after(0, self.restore_from_tray)
        except Exception:
            pass

    def quit_from_tray(self, icon, item):
        try:
            self.after(0, self.shutdown_app)
        except Exception:
            self.shutdown_app()

    def shutdown_app(self):
        self.app_running = False
        try:
            if self._anim_after_id:
                self.after_cancel(self._anim_after_id)
        except Exception:
            pass

        if self.tray_icon:
            try: self.tray_icon.stop()
            except Exception: pass
        self.tray_icon = None
        self._tray_running = False

        if self.arduino and self.arduino.is_open:
            try:
                self.arduino.write(b"{KNOB_CMD_STOP}\n")
            except serial.SerialException:
                pass
            finally:
                try: self.arduino.close()
                except Exception: pass
        if self.listener_thread:
            self.listener_thread.join(timeout=1)
        try:
            comtypes.CoUninitialize()
        except Exception:
            pass

        # Descargar la fuente privada si se cargó
        if hasattr(self, "_font_path") and self._font_path and os.path.exists(self._font_path):
            unload_private_font(self._font_path)

        release_single_instance_mutex(self._mutex_handle)

        try:
            if self._show_event_handle:
                kernel32.CloseHandle(self._show_event_handle)
        except Exception:
            pass

        self.destroy()

    # ======= Utils =======
    def _log_error(self, msg: str):
        try:
            with open(self.error_log, "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass


if __name__ == "__main__":
    # Instancia única
    mutex_handle, acquired = acquire_single_instance_mutex()
    if not acquired:
        signaled = signal_existing_instance_to_show()
        if not signaled and win32gui:
            try:
                hwnd = win32gui.FindWindow(None, APP_TITLE)
                if hwnd:
                    if win32con: win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    else: win32gui.ShowWindow(hwnd, 9)
                    win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
        sys.exit(0)

    show_event_handle = create_auto_reset_event(SHOW_EVENT_NAME)

    try:
        app = AudioMixerApp(mutex_handle=mutex_handle, show_event_handle=show_event_handle)
        app.mainloop()
    except Exception as e:
        err = "".join(traceback.format_exception(type(e), e, e.__traceback__))
        try:
            with open(os.path.join(os.path.expanduser("~"), "knob_error.log"), "a", encoding="utf-8") as f:
                f.write(err + "\n")
        except Exception:
            pass
        try:
            release_single_instance_mutex(mutex_handle)
            if show_event_handle:
                kernel32.CloseHandle(show_event_handle)
        except Exception:
            pass
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Error al iniciar", "Se produjo un error y la app se cerró.\n\nDetalle guardado en knob_error.log")
        root.destroy()
        raise
