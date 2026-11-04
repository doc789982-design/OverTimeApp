#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Уведомление появляется снизу справа, над панелью задач.

Проверяется на настоящем QML-окне: после showCustom нижний и правый
края плашки стоят в доступной области экрана (availableGeometry —
та самая область без панели задач Windows), с отступом 20 px.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_desktop_notification_position.py
"""
import os
import sys
import time
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from PySide6.QtCore import QMetaObject, QObject, Property, Q_ARG, Qt  # noqa: E402
from PySide6.QtQuick import QQuickWindow  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402


WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 900; height: 700; visible: true
    AppUI.AppDesktopNotification { id: note }
}
"""


class StubBackend(QObject):
    isDarkTheme = Property(bool, lambda self: True, constant=True)
    reducedMotion = Property(bool, lambda self: False, constant=True)


def main():
    from PySide6.QtQml import QQmlApplicationEngine

    app = QApplication.instance() or QApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = ROOT / "_test_notification_position.qml"
    wrapper.write_text(WRAPPER, encoding="utf-8")

    def settle(seconds):
        end = time.time() + seconds
        while time.time() < end:
            app.processEvents()
            time.sleep(0.01)

    try:
        engine.load(str(wrapper))
        assert engine.rootObjects(), "окно не загрузилось"
        settle(0.5)
        notes = []
        for win in app.allWindows():
            if not isinstance(win, QQuickWindow):
                continue
            try:
                if win.property("message") is not None and int(win.width()) == 400:
                    notes.append(win)
            except Exception:
                pass
        assert notes, "окно уведомления не найдено"
        note = notes[0]
        assert QMetaObject.invokeMethod(
            note, "showCustom", Qt.ConnectionType.DirectConnection,
            Q_ARG("QVariant", "Проверка положения уведомления"),
            Q_ARG("QVariant", ""))
        settle(0.9)  # анимация въезда справа — 500 мс

        avail = app.primaryScreen().availableGeometry()
        right = int(note.x() + note.width())
        bottom = int(note.y() + note.height())
        wanted_right = int(avail.x() + avail.width() - 20)
        wanted_bottom = int(avail.y() + avail.height() - 20)
        assert abs(right - wanted_right) <= 2, (right, wanted_right)
        assert abs(bottom - wanted_bottom) <= 2, (bottom, wanted_bottom)
        assert note.y() > avail.y() + avail.height() / 2, note.y()
        print("PASS: уведомление стоит снизу справа над панелью задач")
        return 0
    finally:
        try:
            wrapper.unlink()
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
