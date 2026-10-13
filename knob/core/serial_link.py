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
                    with self._latest_lock:
                        self._latest = None
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
            try:
                ok = handshake(ser)
            except Exception as e:
                log(f"{dev}: error en handshake: {e}")
                ok = False
            if ok:
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
