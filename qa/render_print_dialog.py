#!/usr/bin/env python3
"""Рендер окна печати в песочнице: PNG-снимки всех состояний.

Проверяет внешний вид AppPrintDialog без Windows: светлая/тёмная тема,
выбранный пункт «Экспортировать в Excel», режим диапазона страниц.
Снимки кладутся в qa/render/: print_dialog_{light,dark,excel,range}.png

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen QSG_RASTER_BACKEND=1 \\
    python3 qa/render_print_dialog.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
OUT = os.path.join(ROOT, "qa", "render")

from PySide6.QtCore import QObject, Property, Signal, Slot, QTimer
from PySide6.QtGui import QGuiApplication, QImage
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem

PRINTERS = ["Microsoft Print to PDF", "HP LaserJet 1020",
            "\\\\SRV\\Xerox WorkCentre 5225"]


class StubBackend(QObject):
    """Минимальный backend для окна печати."""

    isDarkChanged = Signal()

    def __init__(self):
        super().__init__()
        self._dark = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        if self._dark != bool(v):
            self._dark = bool(v)
            self.isDarkChanged.emit()

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

    @Property(list, constant=True)
    def printerList(self):
        return PRINTERS

    @Property(str, constant=True)
    def defaultPrinter(self):
        return "HP LaserJet 1020"

    @Slot()
    def loadPrinters(self):
        pass


# Обёртка живёт в КОРНЕ репозитория и импортирует компоненты так же,
# как main.qml («import "components" as AppUI») — иначе синглтон AppTheme
# не увидит контекстное свойство backend и тема/списки отвалятся.
WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI

ApplicationWindow {
    id: w
    width: 920; height: 660; visible: true
    color: AppUI.AppTheme.isDark ? "#121212" : "#E4E7EA"
    AppUI.AppPrintDialog { id: dlg }
    // открываем по таймеру: мгновенное открытие в Component.onCompleted
    // ловит биндинги до готовности (кнопка выглядит выключенной)
    Timer { interval: 400; running: true; onTriggered: dlg.open() }
}
"""


def find_by_class(item, klass, acc):
    if item.metaObject().className().startswith(klass):
        acc.append(item)
    for c in item.childItems():
        find_by_class(c, klass, acc)
    return acc


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    # ВАЖНО: держим ссылку на контекст — иначе PySide6 соберёт обёртку
    # и свойство backend в QML станет null.
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_dialog.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось (ошибка QML выше)"
        win = engine.rootObjects()[0]
        # ждём открытия попапа (Timer в обёртке) и появления содержимого
        import time as _t
        deadline = _t.time() + 3
        content = combos = None
        while _t.time() < deadline:
            for _ in range(5):
                app.processEvents()
            top = [i for i in win.findChildren(QQuickItem)
                   if i.parentItem() is None]
            if top:
                found = find_by_class(top[0], "AppComboBox", [])
                if len(found) >= 4:
                    content, combos = top[0], found
                    break
            _t.sleep(0.05)
        assert combos, "не дождались содержимого окна печати"
        printer_combo, range_combo = combos[0], combos[1]
        shots = []

        import time

        def snap(name):
            # дать анимациям открытия/переключения реально закончиться
            deadline = time.time() + 1.5
            while time.time() < deadline:
                app.processEvents()
                time.sleep(0.02)
            img = win.grabWindow()
            path = os.path.join(OUT, "print_dialog_%s.png" % name)
            img.save(path)
            shots.append(path)

        app.processEvents()
        snap("light")

        # режим диапазона страниц
        range_combo.setProperty("currentIndex", 1)
        snap("range")

        # пункт «Экспортировать в Excel» — последний в списке принтеров
        n = printer_combo.property("count")
        assert n == len(PRINTERS) + 1, "в списке %d, ожидали %d (+Excel)" % (
            n, len(PRINTERS))
        model = printer_combo.property("model")
        assert model[-1] == "Экспортировать в Excel", model[-1]
        printer_combo.setProperty("currentIndex", n - 1)
        snap("excel")

        # тёмная тема
        stub.setDark(True)
        printer_combo.setProperty("currentIndex", 1)
        range_combo.setProperty("currentIndex", 0)
        snap("dark")

        for p in shots:
            assert os.path.getsize(p) > 5000, "пустой снимок: %s" % p
        print("снимки: " + ", ".join(os.path.basename(s) for s in shots))
        print("═══ ОКНО ПЕЧАТИ ОТРЕНДЕРЕНО ═══")
        return 0
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
