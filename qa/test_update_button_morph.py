#!/usr/bin/env python3
"""Тест единого морф-значка кнопки обновления (сборки 310, 312).

Раньше кнопка держала четыре отдельные картинки и складывала их
в вертикальную линию при смене состояния. Пользователь попросил
значок, который перетекает между состояниями едиными элементами
и выглядит дорого: теперь фигура одна (дуга + три штриха из
общего узла), а состояния — её позы.

Проверяется:
  A. Позы следуют за бэкендом: refresh → download → ring →
     install → refresh;
  B. Элементы те же, геометрия разная: в покое третий штрих
     схлопнут и дуга открыта; в «скачать» дуга опущена лотком и
     маячок включён; кольцо честно показывает проценты, стрелка
     при загрузке и в «готово» серая (312); в «готово» дуга
     полна, углового маячка нет — волну пускает само кольцо,
     средняя линия волны точно по окружности кольца и с общим
     центром (отзыв на 313-ю);
  C. Проверка (updateChecking) вращает значок, по окончании
     вращение останавливается ровно в исходной позе;
  D. «Меньше анимаций»: вращения нет, позы меняются мгновенно.
     Заодно: у AppTheme появилось свойство reducedMotion, без
     которого тумблер не действовал на компоненты с чтением
     AppTheme.reducedMotion (включая маячок этой кнопки);
  E. Кнопка НЕ берёт клавиатурный фокус: таб-обводки у кнопок
     убраны из всей программы (312). Клик мышью работает.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen QSG_RASTER_BACKEND=1 \
    python3 qa/test_update_button_morph.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from PySide6.QtCore import QObject, Property, Signal, Slot, Qt, QPointF
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem, QQuickWindow  # noqa: F401
from PySide6.QtTest import QTest


class StubBackend(QObject):
    isDarkChanged = Signal()
    remoteUpdateAvailableChanged = Signal()
    remoteDownloadingChanged = Signal()
    remoteDownloadProgressChanged = Signal()
    updateReadyChanged = Signal()
    updateCheckingChanged = Signal()
    reducedMotionChanged = Signal()

    FIELD = {"remoteUpdateAvailable": "_avail",
             "remoteDownloading": "_down",
             "remoteDownloadProgress": "_prog",
             "updateReady": "_ready",
             "updateChecking": "_chk",
             "reducedMotion": "_reduced"}

    def __init__(self):
        super().__init__()
        self._dark = False
        self._avail = self._down = self._ready = self._chk = False
        self._reduced = False
        self._prog = 0
        self.calls = []

    isDarkTheme = Property(bool, lambda s: s._dark, notify=isDarkChanged)
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

    def _set(self, prop, val):
        setattr(self, self.FIELD[prop], val)
        getattr(self, prop + "Changed").emit()

    @Slot()
    def checkAllUpdateSources(self):
        self.calls.append("check")

    @Slot()
    def startRemoteDownload(self):
        self.calls.append("download")

    @Slot()
    def applyReadyUpdate(self):
        self.calls.append("apply")


WIN_QML = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 140; height: 120; visible: true
    Row {
        anchors.centerIn: parent
        AppUI.UpdateDownloadButton { height: 46 }
    }
}
"""


def main() -> int:
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    stub = StubBackend()
    engine.rootContext().setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_test_upd_morph.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WIN_QML)

    def wait_ms(ms):
        t = time.time() + ms / 1000.0
        while time.time() < t:
            app.processEvents()
            time.sleep(0.01)

    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "кнопка не загрузилась"
        win = engine.rootObjects()[0]
        btn = win.findChild(QQuickItem, "updButton")
        glyph = win.findChild(QQuickItem, "updGlyph")
        assert btn and glyph, "updButton/updGlyph не найдены"
        radar = win.findChild(QQuickItem, "updRadar")
        halo = win.findChild(QQuickItem, "updHalo")
        assert radar is not None and halo is not None, \
            "маячок/волна готовности не найдены"
        spinner = glyph.parentItem()
        wait_ms(300)

        # ── A+B. Позы и геометрия ──
        assert btn.property("vis") == "refresh", btn.property("vis")
        assert abs(btn.property("cAlpha")) < 0.05, "в покое штрих C не схлопнут"
        assert abs(btn.property("arcSweep") - 270) < 2, "в покое дуга не открыта"

        stub._set("remoteUpdateAvailable", True)
        wait_ms(800)
        assert btn.property("vis") == "download", btn.property("vis")
        assert btn.property("arcCY") > 13.0, "дуга не опустилась лотком"
        assert abs(btn.property("cAlpha") - 1) < 0.05, "правая грань стрелки не выросла"
        assert radar.property("active"), "в «доступно» угловой маячок выключен"

        stub._set("remoteDownloading", True)
        stub._set("remoteDownloadProgress", 50)
        wait_ms(800)
        assert btn.property("vis") == "ring", btn.property("vis")
        assert abs(btn.property("arcSweep") - 180) < 8, \
            "кольцо не показывает 50%%: %s" % btn.property("arcSweep")
        assert btn.property("segColor").name() == "#5f6b74", \
            "при загрузке стрелка не серая: %s" % btn.property("segColor").name()

        stub._set("remoteDownloading", False)
        stub._set("updateReady", True)
        wait_ms(800)
        assert btn.property("vis") == "install", btn.property("vis")
        assert abs(btn.property("arcSweep") - 360) < 2, "кольцо не дозалито"
        assert btn.property("segColor").name() == "#5f6b74", \
            "в «готово» стрелка не серая: %s" % btn.property("segColor").name()
        assert not radar.property("active"), \
            "в «готово» угловой маячок включён (волну должно пускать кольцо)"

        # Волна готовности ровно по окружности кольца (отзыв на 313-ю:
        # «кружок радара немного криво относительно окружности»):
        # ширина = диаметр средней линии кольца + обводка, центры — один.
        ring_r = float(btn.property("ringStrokeRadiusPx"))
        assert ring_r > 1, "радиус средней линии кольца не посчитался"
        assert abs(float(halo.width()) - (2 * ring_r + 1.6)) < 0.6, \
            "волна не по средней линии кольца: ширина %s при радиусе %s" % \
            (halo.width(), ring_r)
        host = halo.parentItem()
        assert host == glyph.parentItem(), \
            "волна не внутри значка — при наведении отстанет от кольца"
        assert abs((float(halo.x()) + float(halo.width()) / 2)
                   - float(host.width()) / 2) < 0.05, "центр волны смещён по X"
        assert abs((float(halo.y()) + float(halo.height()) / 2)
                   - float(host.height()) / 2) < 0.05, "центр волны смещён по Y"

        stub._set("updateReady", False)
        stub._set("remoteUpdateAvailable", False)
        wait_ms(800)
        assert btn.property("vis") == "refresh", btn.property("vis")
        print("A+B: позы и геометрия единого значка следуют за состояниями ✓")

        # ── C. Проверка вращает и ровно останавливается ──
        stub._set("updateChecking", True)
        wait_ms(300)
        r1 = float(spinner.property("rotation"))
        wait_ms(300)
        r2 = float(spinner.property("rotation"))
        assert r2 > r1, "при проверке значок не вращается (%s -> %s)" % (r1, r2)
        stub._set("updateChecking", False)
        wait_ms(800)
        r3 = float(spinner.property("rotation")) % 360
        assert r3 < 0.5 or r3 > 359.5, "вращение не вернулось в исходную позу: %s" % r3
        print("C: проверка вращает, остановка ровно в исходной позе ✓")

        # ── D. Меньше анимаций ──
        theme = open(os.path.join(ROOT, "components", "AppTheme.qml"),
                     encoding="utf-8").read()
        assert "property bool reducedMotion: backend.reducedMotion" in theme, \
            "у AppTheme нет reducedMotion — тумблер не дойдёт до компонентов"
        stub._set("reducedMotion", True)
        wait_ms(100)
        assert btn.property("reduced"), "кнопка не видит «меньше анимаций»"
        stub._set("updateChecking", True)
        wait_ms(450)
        assert abs(float(spinner.property("rotation")) % 360) < 0.5, \
            "при «меньше анимаций» значок вращается"
        stub._set("updateChecking", False)
        stub._set("remoteUpdateAvailable", True)
        wait_ms(250)   # антидребезг 80 мс + мгновенная поза (анимация заняла бы 540)
        assert abs(btn.property("jointY") - 15.8) < 0.3, \
            "поза при «меньше анимаций» не применилась мгновенно: %s" % btn.property("jointY")
        stub._set("remoteUpdateAvailable", False)
        stub._set("reducedMotion", False)
        wait_ms(300)
        print("D: «меньше анимаций» — без вращения, позы мгновенно; тема прокинута ✓")

        # ── E. Без таб-фокуса; клик мышью ──
        assert btn.property("activeFocusOnTab") is False, \
            "кнопка снова в таб-цепочке"
        pos = btn.mapToScene(QPointF(23, 23)).toPoint()
        QTest.mouseClick(win, Qt.MouseButton.LeftButton,
                         Qt.KeyboardModifier.NoModifier, pos)
        wait_ms(150)
        assert stub.calls == ["check"], stub.calls
        assert not btn.property("activeFocus"), "клик дал кнопке фокус"
        print("E: таб-фокуса нет, клик запускает проверку ✓")
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass

    print("═══ МОРФ-ЗНАЧОК КНОПКИ ОБНОВЛЕНИЯ: ПОЗЫ, ВРАЩЕНИЕ, ЦВЕТА СОСТОЯНИЙ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
