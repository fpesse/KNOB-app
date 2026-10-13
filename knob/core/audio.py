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
        self._meter_ok = {}       # clave -> último intento ok (para no loguear cada refresco)
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
        self._master_meter = self._activate_meter(AudioUtilities.GetSpeakers, "master", self._master_meter)
        self._mic_meter = self._activate_meter(AudioUtilities.GetMicrophone, "mic", self._mic_meter)

    def _activate_meter(self, getter, key: str, current):
        """Medidor del dispositivo; registra solo al pasar a fallar (o el primer fallo)."""
        try:
            dev = getter()
            if not dev:
                raise RuntimeError("sin dispositivo")
            dev = getattr(dev, "_dev", dev)      # pycaw nuevo: AudioDevice envuelve al IMMDevice
            meter = dev.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None
                                 ).QueryInterface(IAudioMeterInformation)
            self._meter_ok[key] = True
            return meter
        except Exception as e:
            if self._meter_ok.get(key, True):
                log(f"audio: medidor {key}: {e}")
            self._meter_ok[key] = False
            return current

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
            mic = getattr(mic, "_dev", mic)
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
