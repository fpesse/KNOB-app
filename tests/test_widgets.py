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
