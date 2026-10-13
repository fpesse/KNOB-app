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
