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
