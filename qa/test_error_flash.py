#!/usr/bin/env python3
"""Красная рамка ошибки-ВСПЫШКА (AppTextField / AppDateField).

Рамка ошибки по всему приложению ведёт себя как вспышка:
  – flashError() зажигает рамку на 3 секунды;
  – затем она ПЛАВНО гаснет (не висит, пока не введёшь значение);
  – повторное «Сохранить» зажигает вспышку заново — даже пока
    предыдущая ещё горит (таймер перезапускается);
  – ввод значения гасит рамку сразу.

Проверяется на изолированных полях и на окне ведомости
(пустые номер и дата приказа → оба поля вспыхивают).

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_error_flash.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 900; height: 400; visible: true
    Column {
        spacing: 16; x: 20; y: 20
        AppUI.AppTextField { id: tf; objectName: "tf"; width: 300; label: "Поле" }
        AppUI.AppDateField { id: df; objectName: "df"; width: 300; label: "Дата" }
    }
    AppUI.MoneyOrderDialog { id: dlg }
    Component.onCompleted: dlg.openNew()
}
"""

EMPS = [
    {"id": 1, "name": "Иванов Иван", "subtitle": "капитан",
     "hours": 8, "overtime": 4, "days": 2},
]


from PySide6.QtCore import QObject, Slot, QMetaObject
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine, QJSValue
from PySide6.QtQuick import QQuickItem   # конвертер QQuickItem*
from PySide6.QtTest import QTest
from PySide6.QtCore import Qt


class StubBackend(QObject):
    @Slot(result="QVariant")
    def moneyOrderEmployees(self):
        return EMPS

    @Slot(str, str, str, str, result="QVariant")
    def saveMoneyOrder(self, payload, no, dt, comment):
        return {"ok": True, "count": 1}


def to_var(v):
    return v.toVariant() if isinstance(v, QJSValue) else v


def main() -> int:
    app = QGuiApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()              # ссылка обязательна (GC)
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_flash.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)

    errors = []
    engine.warnings.connect(
        lambda ws: errors.extend(str(x.toString()) for x in ws))
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]
        deadline = time.time() + 3
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.02)

        tf = df = None
        for o in win.findChildren(QObject):
            if o.objectName() == "tf":
                tf = o
            elif o.objectName() == "df":
                df = o
        assert tf is not None and df is not None, "поля не найдены"
        dlg = None
        for o in win.findChildren(QObject):
            if o.metaObject().className().startswith("MoneyOrderDialog"):
                dlg = o
                break
        assert dlg is not None, "диалог ведомости не найден"

        def settle(s, n=25):
            # прокачиваем события ВО ВРЕМЯ ожидания: анимации Qt
            # тикают только на processEvents (sleep без прокачки
            # останавливает вспышку и затухание)
            end = time.time() + s
            while time.time() < end:
                app.processEvents()
                time.sleep(0.01)
            for _ in range(n):
                app.processEvents()

        def glow(field):
            return float(field.property("errGlow") or 0)

        # ── 1. вспышка горит ──
        QMetaObject.invokeMethod(tf, "flashError")
        settle(0.4)
        assert tf.property("hasError") is True
        assert glow(tf) > 0.9, "рамка не зажглась: %s" % glow(tf)
        print("1. flashError(): рамка зажглась (errGlow=%.2f) ✓" % glow(tf))

        # ── 2. ввод гасит сразу ──
        tf.forceActiveFocus()
        QTest.keyClick(win, Qt.Key_4)     # QWindow-вариант: ввод в сфокусированное поле
        settle(0.9)
        assert tf.property("hasError") is False
        assert glow(tf) < 0.05, "ввод не погасил рамку: %s" % glow(tf)
        print("2. Ввод значения погасил рамку сразу (errGlow=%.2f) ✓" % glow(tf))

        # ── 3. три секунды горит, потом плавно гаснет ──
        QMetaObject.invokeMethod(tf, "flashError")
        settle(0.4)
        settle(2.2)
        assert glow(tf) > 0.9, "рамка погасла раньше трёх секунд: %s" % glow(tf)
        print("3а. Рамка горит дольше двух секунд (errGlow=%.2f) ✓" % glow(tf))
        settle(1.9)      # 3с истекли + fade
        assert tf.property("hasError") is False, "hasError не сбросился таймером"
        assert glow(tf) < 0.05, "рамка не погасла после 3 секунд: %s" % glow(tf)
        print("3б. Через 3 секунды рамка плавно погасла (errGlow=%.2f) ✓" % glow(tf))

        # ── 4. повторная вспышка ПОКА ГОРИТ: таймер перезапускается ──
        QMetaObject.invokeMethod(tf, "flashError")
        settle(1.0)
        assert glow(tf) > 0.9
        QMetaObject.invokeMethod(tf, "flashError")    # снова «Сохранить»
        settle(2.5)      # с первой вспышки прошло 3.5с — без перезапуска уже погасла бы
        assert glow(tf) > 0.9, \
            "повторная вспышка не перезапустила таймер: %s" % glow(tf)
        print("4. Повторное сохранение перезапустило вспышку (errGlow=%.2f) ✓"
              % glow(tf))
        settle(1.6)
        assert glow(tf) < 0.05

        # ── 5. поле даты: та же вспышка, ввод даты гасит ──
        QMetaObject.invokeMethod(df, "flashError")
        settle(0.4)
        assert glow(df) > 0.9, "у поля даты рамка не зажглась"
        df.setProperty("selectedDate", "2026-10-12")
        settle(0.9)
        assert glow(df) < 0.05, "ввод даты не погасил рамку: %s" % glow(df)
        print("5. Поле даты: вспышка зажглась и погасла от ввода даты ✓")

        # ── 6. ведомость: пустые № и дата → оба поля вспыхнули,
        #      затухли сами, повторное сохранение — снова вспышка ──
        dlg.openNew()
        settle(0.4)
        fields = [o for o in dlg.findChildren(QObject)
                  if o.metaObject().className().startswith("AppTextField")
                  and o.property("label") == "№ приказа"]
        dates = [o for o in dlg.findChildren(QObject)
                 if o.metaObject().className().startswith("AppDateField")]
        assert fields and dates
        no_f, date_f = fields[0], dates[0]
        # дата предзаполнена сегодняшней — очищаем, как руками
        date_f.setProperty("text", "")
        settle(0.2)
        QMetaObject.invokeMethod(dlg, "accepted")
        settle(0.4)
        assert no_f.property("hasError") is True and glow(no_f) > 0.9
        assert date_f.property("hasError") is True and glow(date_f) > 0.9
        print("6а. Ведомость без № и даты: оба поля вспыхнули ✓")
        settle(3.9)
        assert no_f.property("hasError") is False and glow(no_f) < 0.05
        assert date_f.property("hasError") is False and glow(date_f) < 0.05
        print("6б. Обе рамки сами погасли через 3 секунды ✓")
        QMetaObject.invokeMethod(dlg, "accepted")   # снова «Провести приказ»
        settle(0.4)
        assert no_f.property("hasError") is True and glow(no_f) > 0.9
        assert date_f.property("hasError") is True and glow(date_f) > 0.9
        print("6в. Повторное «Провести приказ»: вспышка снова ✓")

        bad = [e for e in errors
               if "not a function" in e or "Cannot assign" in e
               or "read-only" in e or "ReferenceError" in e
               or ("TypeError" in e and "isDarkTheme" not in e)]
        assert not bad, "ошибки QML: %r" % (bad[:3],)

        print("═══ ВСПЫШКА ОШИБКИ: 3 СЕКУНДЫ → ПЛАВНО ГАСНЕТ, ПЕРЕЗАПУСК РАБОТАЕТ ═══")
        return 0
    finally:
        if os.path.exists(wrapper):
            os.remove(wrapper)


if __name__ == "__main__":
    sys.exit(main())
