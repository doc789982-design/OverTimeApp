#!/usr/bin/env python3
"""Рендер и проверка левой панели управления (LeftControlPanel).

Панель рендерится со стабом backend; проверяется вёрстка после переноса
кнопки печати в ряд «тема · настройки · печать»:
  – в верхнем ряду у правого края три иконки (тема, настройки, печать);
  – в нижнем ряду широкая кнопка «Новый сотрудник» и круглая
    кнопка-₽ (Приказ всем) справа от неё — только символ рубля;
  – иконок во втором ряду (между рядами) нет.
Снимок: qa/render/panel.png

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \\
    QSG_RASTER_BACKEND=1 python3 qa/test_panel_view.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "qa", "render")

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI

ApplicationWindow {
    id: w
    width: 920; height: 300; visible: true
    color: AppUI.AppTheme.isDark ? "#121212" : "#E4E7EA"
    AppUI.LeftControlPanel {
        id: panel
        anchors.top: parent.top
        anchors.left: parent.left
        anchors.right: parent.right
    }
}
"""


from PySide6.QtCore import QObject, Property
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine


class StubBackend(QObject):
    """Стаб на уровне модуля: локальные классы с Property PySide6
    регистрирует ненадёжно (окно рендерится как голый QWindow)."""

    @Property(str, constant=True)
    def activeDepartmentName(self):
        return "Отдел информационных технологий"

    @Property(list, constant=True)
    def dbList(self):
        return []

    @Property(str, constant=True)
    def currentPeriodText(self):
        return "Сентябрь 2026"

    @Property(bool, constant=True)
    def isDarkTheme(self):
        return False


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()          # ссылка обязательна (см. render_print_dialog)
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_panel.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "панель не загрузилась"
        # даём окну построиться; берём готовое QQuickWindow из topLevelWindows:
        # обёртка rootObjects()[0], взятая сразу после load, бывает голым QWindow
        # (без grabWindow) — объект ещё не финализирован
        deadline = time.time() + 2.5
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.02)
        from PySide6.QtQuick import QQuickWindow
        win = next(x for x in app.topLevelWindows()
                   if isinstance(x, QQuickWindow))
        img = win.grabWindow()
        shot = os.path.join(OUT, "panel.png")
        img.save(shot)
        assert os.path.getsize(shot) > 8000, "пустой снимок"

        W = img.width()

        def clusters(y0, y1, x0, x1):
            """Число групп тёмных пикселей по горизонтали (иконки/глифы)."""
            cols = []
            for x in range(x0, x1):
                hit = any(img.pixelColor(x, y).red() + img.pixelColor(x, y).green()
                          + img.pixelColor(x, y).blue() < 690
                          for y in range(y0, y1, 2))
                cols.append(hit)
            groups = 0
            prev = False
            for v in cols:
                if v and not prev:
                    groups += 1
                prev = v
            return groups

        # верхний ряд: три иконки (тема · настройки · печать) у правого края.
        # Заголовок подразделения сюда не дотягивается (клипуется до иконок),
        # зона заведомо иконочная: последние ~150 px минус поле карточки.
        top_icons = clusters(32, 58, W - 150, W - 14)
        assert top_icons == 3, "в верхнем ряду %d иконок, ожидали 3 (тема/настройки/печать)" % top_icons

        # широкая кнопка «Новый сотрудник»: карточка (не-панельный фон на y=78,
        # между рядами) и границы кнопки (не-белые внутри карточки на y=84)
        panel_bg = img.pixelColor(2, 78).name()
        xs = [x for x in range(W) if img.pixelColor(x, 78).name() != panel_bg]
        card_left, card_right = min(xs), max(xs)
        xs2 = [x for x in range(card_left + 1, card_right)
               if img.pixelColor(x, 84).name() != "#ffffff"]
        btn_left, btn_right = min(xs2), max(xs2)
        left_pad = btn_left - card_left
        right_pad = card_right - btn_right
        assert left_pad > 0 and right_pad > 0, "кнопка не внутри карточки"
        assert abs(left_pad - right_pad) <= 6, \
            "поля кнопки несимметричны: слева %d, справа %d" % (left_pad, right_pad)

        # объектная проверка нижнего ряда: широкая «Новый сотрудник»
        # + круглая кнопка-₽ (сборка 263)
        from PySide6.QtCore import QMetaObject
        panel = None
        for o in win.findChildren(QObject):
            if o.metaObject().className().startswith("LeftControlPanel"):
                panel = o
                break
        assert panel is not None, "панель не найдена"
        emp_btn = icon_btn = None
        for o in panel.findChildren(QObject):
            if o.metaObject().className().startswith("AppButton") \
                    and str(o.property("text") or "") == "Новый сотрудник":
                emp_btn = o
            if o.metaObject().className().startswith("AppIconButton") \
                    and "ruble.svg" in str(o.property("iconSource") or ""):
                icon_btn = o
        assert emp_btn is not None, "нет кнопки «Новый сотрудник»"
        assert icon_btn is not None, "нет круглой кнопки-₽ (ruble.svg)"
        assert float(emp_btn.property("width")) > 250, \
            "«Новый сотрудник» не расширилась: %s px" % emp_btn.property("width")
        iw, ih = float(icon_btn.property("width")), float(icon_btn.property("height"))
        assert abs(iw - 36) < 1 and abs(ih - 36) < 1, \
            "кнопка-₽ не круглая 36×36: %s×%s" % (iw, ih)
        gap = float(icon_btn.property("x")) - (float(emp_btn.property("x"))
                                                + float(emp_btn.property("width")))
        assert 4 <= gap <= 20, "кнопка-₽ не рядом с «Новым сотрудником»: %s px" % gap
        for o in panel.findChildren(QObject):
            if o.metaObject().className().startswith("AppButton") \
                    and "Приказ всем" in str(o.property("text") or ""):
                raise AssertionError("широкая кнопка «Приказ всем» не убрана")

        print("верхний ряд: 3 иконки (тема · настройки · печать) ✓")
        print("нижний ряд: широкая «Новый сотрудник» %d px + круглая ₽ 36×36 "
              "через %d px, поля %d/%d px ✓"
              % (emp_btn.property("width"), gap, left_pad, right_pad))
        print("снимок: qa/render/panel.png")
        print("═══ ПАНЕЛЬ: ВЁРСТКА ПО ЗАДАЧЕ ═══")
        return 0
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
