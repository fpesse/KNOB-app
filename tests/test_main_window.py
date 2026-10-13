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


def test_perfil_con_ampersand_se_escapa(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.theme import get_theme
    w = MainWindow(get_theme("Nebula"))
    w.set_profiles(["Default", "R&B"], "Default")
    assert w._profile_menu.actions()[1].text().strip().replace("✓", "").strip() == "R&&B"


@pytest.mark.parametrize("cls_name", ["channel_card.ChannelCard", "encoder_card.EncoderCard"])
def test_subtitulo_de_foco(qapp, cls_name):
    import importlib
    from knob.ui.theme import get_theme
    mod, cls = cls_name.split(".")
    card = getattr(importlib.import_module(f"knob.ui.widgets.{mod}"), cls)(0, get_theme("Nebula"))
    card.show()
    card.set_assignment("App en foco", None, "focus")
    assert card.sub_label.text() == "→ ninguna"
    card.set_focus_target("Chrome")
    assert card.sub_label.text() == "→ Chrome"
    card.set_assignment("Discord", None, None)
    assert not card.sub_label.isVisibleTo(card)


def test_app_icons_cache_y_reintento(qapp, monkeypatch):
    import sys
    from knob.ui import icons
    calls = []

    def lookup(name):
        calls.append(name)
        return sys.executable if name.lower() == "python" else None

    ai = icons.AppIcons(lookup)
    p1 = ai.icon_for("Python")
    assert p1 is not None and not p1.isNull()
    assert ai.icon_for("python") is p1
    assert calls == ["Python"]
    clock = {"t": 100.0}
    monkeypatch.setattr(icons.time, "monotonic", lambda: clock["t"])
    assert ai.icon_for("nada") is None
    assert ai.icon_for("nada") is None
    assert calls.count("nada") == 1
    clock["t"] += icons.RETRY_S + 1
    assert ai.icon_for("nada") is None
    assert calls.count("nada") == 2


def test_app_icons_ruta_inexistente(qapp):
    from knob.ui.icons import AppIcons
    assert AppIcons(lambda n: r"C:\no\existe.exe").icon_for("x") is None


def test_output_bar_nombre_y_flechas(qapp):
    from PySide6.QtWidgets import QToolButton
    from knob.ui.theme import get_theme
    from knob.ui.widgets.output_bar import OutputBar
    bar = OutputBar(get_theme("Nebula"))
    bar.set_name(None)
    assert bar.name_label.text() == "Sin dispositivo de salida"
    got = []
    bar.prev_clicked.connect(lambda: got.append("prev"))
    bar.next_clicked.connect(lambda: got.append("next"))
    btns = bar.findChildren(QToolButton)
    assert len(btns) == 2
    btns[0].click()
    btns[1].click()
    assert got == ["prev", "next"]


def test_boton_de_perfil_elide_y_escapa(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.theme import get_theme
    w = MainWindow(get_theme("Nebula"))
    w.set_profiles(["P"], "M" * 40)
    assert "…" in w.profile_btn.text() or "..." in w.profile_btn.text()   # Qt 6.8 elide con tres puntos
    assert w.profile_btn.toolTip() == "M" * 40
    w.set_profiles(["R&B"], "R&B")
    assert "R&&B" in w.profile_btn.text()


def test_barra_superior_sin_solapes(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.theme import build_qss, get_theme
    t = get_theme("Nebula")
    qapp.setStyleSheet(build_qss(t))
    w = MainWindow(t)
    w.set_profiles(["W" * 24], "W" * 24)
    w.status.set_state("busy", "COM12 ocupado", "Cerrá el Serial Monitor de Arduino (COM12 ocupado)")
    w.show()
    qapp.processEvents()
    assert not w.status.geometry().intersects(w.profile_btn.geometry())
    assert not w.brand.geometry().intersects(w.status.geometry())
    assert w.status.toolTip().startswith("Cerrá el Serial Monitor")
    w.hide()
