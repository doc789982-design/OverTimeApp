#!/usr/bin/env python3
"""Проверка видимости ошибки в окне компенсации (песочница).

Сценарий: пересечение дней с существующей компенсацией. Проверка
срабатывала, окно тряслось — но сообщение писалось в самый низ
скроллируемой колонки, за нижний край: человек видел только тряску
«непонятно из-за чего». Теперь в режиме периода ошибка встаёт рядом
с полями периода, а окно дополнительно прокручивается к тексту.

Рендерит DayCompDialog со стаб-бэкендом: период 16.11–31.12.2026,
checkDayConflicts возвращает конфликт, жмём «Сохранить» (сигнал
accepted). Проверяем: сохранение не вызвано, ошибка видима, лежит
в пределах вьюпорта скролла и реально нарисована на снимке окна
(красный текст в карточке).

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \\
    QSG_RASTER_BACKEND=1 python3 qa/test_day_comp_dialog_view.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
OUT = os.path.join(ROOT, "qa", "render")

from PySide6.QtCore import QObject, Property, Signal, Slot, QMetaObject, QTimer
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem
from PySide6.QtCore import QPointF

CONFLICT_MSG = ("День 16.11.2026 уже входит в компенсацию за период "
                "16.11.2026 – 28.01.2027 (47 дн., за пред. год)")


class StubBackend(QObject):
    """Минимальный backend для окна компенсации."""

    isDarkChanged = Signal()

    def __init__(self):
        super().__init__()
        self._dark = False
        self.saved = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        if self._dark != bool(v):
            self._dark = bool(v)
            self.isDarkChanged.emit()

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

    def getShift(self):
        return False

    isSelectedEmployeeShift = Property(bool, getShift, constant=True)

    @Slot(str, str, result="QVariant")
    def getWorkingDaysForPeriod(self, start_iso, end_iso):
        from datetime import date, timedelta
        s = date.fromisoformat(start_iso)
        e = date.fromisoformat(end_iso)
        out, cur = [], s
        while cur <= e and len(out) < 400:
            if cur.weekday() < 5:
                out.append(cur.isoformat())
            cur += timedelta(days=1)
        return out

    @Slot(int, result="QVariant")
    def getAvailableBalances(self, year):
        return {"hours": 128, "overtime": 0, "days": 31}

    @Slot(str, str, str, result="QVariant")
    def getShiftDatesForPeriod(self, start_iso, end_iso):
        return {"error": "", "dates": [], "confidence": -1,
                "cycle_days": 0, "work_days": 0}

    @Slot(str, int, result="QVariant")
    def checkDayConflicts(self, dates_csv, exclude_comp_id):
        return {"has": True, "message": CONFLICT_MSG, "count": 1}

    @Slot(str, str, str, str, bool)
    def saveCompensation(self, dates_csv, comp_type, amount, comment, prev):
        self.saved = True


WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI

ApplicationWindow {
    id: w
    width: 920; height: 660; visible: true
    color: "#E4E7EA"
    AppUI.DayCompDialog { id: dlg }
    Timer { interval: 400; running: true; onTriggered: dlg.open() }
}
"""


def find_by_class(item, klass, acc):
    if item.metaObject().className().startswith(klass):
        acc.append(item)
    for c in item.childItems():
        find_by_class(c, klass, acc)
    return acc


def top_item(win):
    top = [i for i in win.findChildren(QQuickItem) if i.parentItem() is None]
    assert top, "нет корневого айтема"
    return top[0]


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_comp_dialog.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]

        # ждём открытия попапа и появления полей периода
        deadline = time.time() + 5
        while time.time() < deadline:
            for _ in range(5):
                app.processEvents()
            if len(find_by_class(top_item(win), "AppDateField", [])) >= 2:
                break
            time.sleep(0.05)
        root0 = top_item(win)

        # сам диалог — Popup (QObject, не айтем): ищем среди детей окна
        from PySide6.QtCore import QObject as _QObject
        dlgs = [o for o in win.findChildren(_QObject)
                if o.metaObject().className().startswith("DayCompDialog")]
        assert dlgs, "окно компенсации не найдено"
        dlg = dlgs[0]

        # настраиваем: дата 16.11.2026, период 16.11–31.12.2026
        dlg.setProperty("targetDate", "2026-11-16")
        fields = find_by_class(root0, "AppDateField", [])
        assert len(fields) >= 2, "нет полей дат: %d" % len(fields)
        fields[0].setProperty("selectedDate", "2026-11-16")
        for _ in range(5):
            app.processEvents()
        fields[1].setProperty("selectedDate", "2026-12-31")
        for _ in range(5):
            app.processEvents()

        # переключаем в режим «Период» (compCol.compMode = 1)
        for col in find_by_class(root0, "QQuickColumn", []):
            if col.property("compMode") is not None:
                col.setProperty("compMode", 1)
                break
        else:
            raise AssertionError("не нашли колонку с compMode")
        for _ in range(10):
            app.processEvents()
        time.sleep(0.2)
        for _ in range(10):
            app.processEvents()

        # «Сохранить»: эмитим сигнал accepted — работает тот же обработчик
        ok = QMetaObject.invokeMethod(dlg, "accepted")
        assert ok, "не удалось вызвать accepted"
        time.sleep(1.6)          # прокрутка + тряска + рендер
        for _ in range(10):
            app.processEvents()

        # 1. сохранение не вызывалось
        assert not stub.saved, "saveCompensation вызван несмотря на конфликт"

        # 2. ошибка показана — тот самый текст
        texts = [t for t in find_by_class(root0, "QQuickText", [])
                 if "уже входит" in str(t.property("text") or "")]
        assert texts, "текст ошибки не найден"
        msg = texts[0]
        assert msg.isVisible(), "текст ошибки скрыт"
        assert str(msg.property("text")).startswith("День 16.11.2026"), \
            msg.property("text")

        # 3. ошибка в пределах видимой части скролла
        flicks = find_by_class(root0, "QQuickFlickable", [])
        assert flicks, "скролл не найден"
        fl = flicks[0]
        pos = msg.mapToItem(fl, QPointF(0, 0)).y()
        cy = float(fl.property("contentY") or 0)
        vh = float(fl.height())
        h = float(msg.height())
        assert pos + h > cy and pos < cy + vh, \
            "ошибка за краем скролла: pos=%s contentY=%s высота=%s" % (pos, cy, vh)

        # 4. ошибка реально нарисована: красный текст на снимке окна
        img = win.grabWindow()
        assert not img.isNull(), "снимок пустой"
        snap = os.path.join(OUT, "comp_dialog_conflict.png")
        img.save(snap)
        red = 0
        for y in range(0, img.height(), 2):
            for x in range(0, img.width(), 2):
                c = img.pixelColor(x, y)
                if c.red() > 160 and c.green() < 115 and c.blue() < 135 \
                        and c.red() - c.green() > 60:
                    red += 1
        assert red >= 25, "красный текст не найден на снимке (%d пикселей)" % red

        print("конфликт периода: сообщение видно в окне (%d кр. пикселей), "
              "сохранение заблокировано ✓" % red)
        print("═══ ОШИБКА ПЕРЕСЕЧЕНИЯ ДНЕЙ ВИДНА В ОКНЕ КОМПЕНСАЦИИ ═══")
        return 0
    finally:
        if os.path.exists(wrapper):
            os.remove(wrapper)


if __name__ == "__main__":
    raise SystemExit(main())
