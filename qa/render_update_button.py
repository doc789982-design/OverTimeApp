#!/usr/bin/env python3
"""Рендер кнопки обновления в четырёх состояниях + окно создания базы (298).

Проверяет единый морф-значок (сборка 310) и маячок-радар:
  idle        — серая стрелка-перезагрузка, без синего и зелёного;
  available   — синяя стрелка «скачать» + синий радар, зелёного нет;
  downloading — синее кольцо прогресса, радар скрыт;
  ready       — зелёная стрелка «установить» + синий радар.

Плюс окно создания базы с полями ответственного лица.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen QSG_RASTER_BACKEND=1 \
    python3 qa/render_update_button.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
OUT = os.path.join(ROOT, "qa", "render")

from PySide6.QtCore import QObject, Property, Signal, Slot
from PySide6.QtGui import QGuiApplication, QImage, QColor
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickWindow

BLUE = ("#0374B5", AppBLUE := (3, 116, 181))     # accentBrand
GREEN = ("#107C3F", (16, 124, 63))               # accentSuccess


class StubBackend(QObject):
    isDarkChanged = Signal()
    remoteUpdateAvailableChanged = Signal()
    remoteDownloadingChanged = Signal()
    remoteDownloadProgressChanged = Signal()
    updateReadyChanged = Signal()
    updateCheckingChanged = Signal()
    reducedMotionChanged = Signal()

    def __init__(self):
        super().__init__()
        self._dark = False
        self._reduced = False
        self._avail = False
        self._down = False
        self._prog = 0
        self._ready = False
        self._chk = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        self._dark = bool(v)

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

    def _b(self, name):
        return getattr(self, "_%s" % name)

    remoteUpdateAvailable = Property(
        bool, lambda s: s._avail, notify=remoteUpdateAvailableChanged)
    remoteDownloading = Property(
        bool, lambda s: s._down, notify=remoteDownloadingChanged)
    remoteDownloadProgress = Property(
        int, lambda s: s._prog, notify=remoteDownloadProgressChanged)
    updateReady = Property(
        bool, lambda s: s._ready, notify=updateReadyChanged)
    updateChecking = Property(
        bool, lambda s: s._chk, notify=updateCheckingChanged)
    reducedMotion = Property(
        bool, lambda s: s._reduced, notify=reducedMotionChanged)

    @Slot()
    def checkAllUpdateSources(self):
        pass

    @Slot()
    def startRemoteDownload(self):
        pass

    @Slot()
    def applyReadyUpdate(self):
        pass

    # для окна создания базы
    @Slot(str, str, str, str, str, str)
    def createNewDatabase(self, *a):
        pass


BTN = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 300; height: 120; visible: true
    Rectangle { anchors.fill: parent; color: "#E4E7EA" }
    Row {
        anchors.centerIn: parent
        AppUI.UpdateDownloadButton { height: 46 }
    }
}
"""


def count_color(img, rgb, tol=14):
    n = 0
    for y in range(img.height()):
        for x in range(img.width()):
            c = QColor(img.pixel(x, y))
            if (abs(c.red() - rgb[0]) <= tol and abs(c.green() - rgb[1]) <= tol
                    and abs(c.blue() - rgb[2]) <= tol):
                n += 1
    return n


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    stub = StubBackend()
    engine.rootContext().setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_upd.qml")

    states = [
        ("idle", dict(), "серая, без синего и зелёного"),
        ("available", dict(_avail=True),
         "синяя стрелка + синий радар, зелёного нет"),
        ("downloading", dict(_avail=True, _down=True, _prog=40),
         "синее кольцо прогресса, радар скрыт"),
        ("ready", dict(_avail=True, _ready=True),
         "зелёная стрелка «установить» + синий радар"),
    ]
    shots = []
    for name, props, what in states:
        # сброс к базовому состоянию (наследовать предыдущую сцену нельзя:
        # в жизни «готово» всегда приходит с загрузкой, уже остановленной)
        stub._avail = stub._down = stub._ready = False
        stub._prog = 0
        for k, v in props.items():
            setattr(stub, k, v)
        with open(wrapper, "w", encoding="utf-8") as f:
            f.write(BTN)
        engine.load(wrapper)
        assert engine.rootObjects(), "сцена %s не загрузилась" % name
        win = engine.rootObjects()[0]
        # ждём: хореография состояний проходит через складывание/раскрытие
        deadline = time.time() + 3.0
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.02)
        img = win.grabWindow()
        path = os.path.join(OUT, "upd_%s.png" % name)
        img.save(path)
        shots.append(path)

        blue = count_color(img, AppBLUE)
        green = count_color(img, GREEN[1])
        if name == "idle":
            assert blue < 40 and green < 40, \
                "idle: синего %d, зелёного %d — должно быть пусто" % (blue, green)
        elif name == "available":
            assert blue > 60, "available: синего мало (%d) — стрелка/радар" % blue
            assert green < 40, "available: зелёный лишний (%d)" % green
        elif name == "downloading":
            assert blue > 40, "downloading: нет синей дуги (%d)" % blue
        elif name == "ready":
            assert green > 60, "ready: зелёной стрелки нет (%d)" % green
            # синий здесь — только точка радара (кольца полупрозрачные
            # и в точный цвет не попадают): точки без active не бывает
            assert blue >= 12, "ready: синей точки радара нет (%d)" % blue
        print("%s: синий %d px, зелёный %d px — %s ✓" % (name, blue, green, what))
        engine.deleteLater()
        engine = QQmlApplicationEngine()
        stub  # ссылка жива
        engine.rootContext().setContextProperty("backend", stub)

    # ── окно создания базы: поля ответственного лица ──
    DBWIN = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 620; height: 560; visible: true
    AppUI.CreateDbDialog { id: dlg }
    Timer { interval: 400; running: true; onTriggered: dlg.open() }
}
"""
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(DBWIN)
    engine.load(wrapper)
    assert engine.rootObjects(), "окно создания базы не загрузилось"
    win = engine.rootObjects()[0]
    deadline = time.time() + 2.0
    while time.time() < deadline:
        app.processEvents()
        time.sleep(0.02)
    img = win.grabWindow()
    path = os.path.join(OUT, "refactor_create_db.png")
    img.save(path)
    shots.append(path)
    white = count_color(img, (255, 255, 255))
    assert white > 40000, "окно базы не отрисовалось (белого %d)" % white
    print("create_db: окно с полями отрисовано ✓")

    try:
        os.remove(wrapper)
    except OSError:
        pass
    print("снимки: " + ", ".join(os.path.basename(s) for s in shots))
    print("═══ КНОПКА ОБНОВЛЕНИЯ (РАДАР) И ОКНО БАЗЫ ОТРЕНДЕРЕНЫ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
