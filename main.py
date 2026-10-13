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
