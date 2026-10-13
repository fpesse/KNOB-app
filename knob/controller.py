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
        self._quitting = False
        self._vu_peak = [0.0, 0.0, 0.0]
        self._connect_signals()
        self._timers = []
        for interval, slot in ((16, self._tick), (33, self._update_vu),
                               (300, self._update_focus_labels), (10000, self._refresh_cards)):
            timer = QTimer(self)
            timer.setInterval(interval)
            timer.timeout.connect(slot)
            self._timers.append(timer)

    # --- persistencia tolerante a fallos -----------------------------------------------
    @staticmethod
    def _persist(fn, *args):
        """Llama a una operación de Config que guarda en disco. Si falla el guardado (OSError)
        se registra y el estado en memoria se conserva. Devuelve (ok, resultado)."""
        try:
            return True, fn(*args)
        except OSError as e:
            log(f"no se pudo guardar la configuración: {e}")
            return False, None

    def _set_config_attr(self, name: str, value) -> None:
        try:
            setattr(self.config, name, value)
        except OSError as e:
            log(f"no se pudo guardar la configuración ({name}): {e}")

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
        threading.Thread(target=self._load_outputs, daemon=True).start()
        threading.Thread(target=self._check_update, daemon=True).start()
        for timer in self._timers:
            timer.start()
        if self._start_minimized and Tray.available():
            self.hide_to_tray()
        else:
            self.window.show()

    def _load_outputs(self) -> None:
        name = ""
        try:
            name = self.outputs.load() or ""
        except Exception as e:
            log(f"outputs.load: {e}")
        self._output_loaded.emit(name)

    def _on_close_requested(self) -> None:
        if self._quitting:
            return
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
        if self._quitting:
            return
        self._quitting = True
        for timer in self._timers:
            timer.stop()
        self.serial.stop()
        self.tray.hide()
        self.window.overlay.close_overlay()
        self.window.hide()
        self.app.quit()

    # --- Effects (los llama el motor; nunca deben lanzar) ---------------------------
    def set_volume(self, target: str, level: float) -> None:
        try:
            self.audio.set_volume(target, level)
        except Exception:
            pass        # sesión muerta o dispositivo desconectado: se reintenta en el próximo tick

    def media_key(self, action: str) -> None:
        try:
            system.send_media_key(action)
        except Exception as e:
            log(f"media_key: {e}")

    def send_hotkey(self, combo: str) -> None:
        def run():
            try:
                system.send_hotkey(combo)
            except Exception as e:
                log(f"send_hotkey: {e}")

        threading.Thread(target=run, daemon=True).start()

    def step_master(self, delta: float) -> None:
        try:
            self.audio.step_master(delta)
        except Exception as e:
            log(f"step_master: {e}")

    def cycle_output(self, direction: int) -> None:
        try:
            name = self.outputs.cycle(direction)
            if name:
                self.window.output_bar.set_name(name, flash=True)
        except Exception as e:
            log(f"cycle_output: {e}")

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
            self._persist(self.config.remember_exe_path, app_name, path)
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
        self._set_config_attr("last_port", port)

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
        self._persist(self.config.set_channel_assignment, ch, value)
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
        self._persist(self.config.set_encoder_assignments, self.engine.encoder_assignments)
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
                    self._persist(self.config.create_profile, name, self.engine.channel_assignments,
                                  self.engine.encoder_assignments)
                else:
                    self._persist(self.config.rename_profile, active, name)
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
            self._persist(self.config.switch_profile, name)
        except ProfileError:
            return
        self._apply_active_profile()

    def _delete_profile(self) -> None:
        before = self.config.profile_names()
        self._persist(self.config.delete_profile, self.config.active_profile)
        if self.config.profile_names() != before:      # también si falló el guardado: la memoria ya cambió
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
        self._set_config_attr("theme", name)
        self.theme = get_theme(name)
        self.app.setStyleSheet(build_qss(self.theme))
        self.window.apply_theme(self.theme)
        if self._settings is not None:
            self._settings.set_theme(self.theme)

    def _set_tray(self, on: bool) -> None:
        self._set_config_attr("minimize_to_tray", on)

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
        try:
            updater.launch_installer(path)
        except Exception as e:
            log(f"lanzar instalador: {e}")
            self._set_update_state("error", "No se pudo abrir el instalador")
            return
        QTimer.singleShot(500, self.quit)     # salir de verdad para que el instalador pueda reemplazar archivos
