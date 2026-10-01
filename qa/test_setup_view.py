#!/usr/bin/env python3
"""Тест окна мастера установки (components/SetupWizard.qml).

Мастер грузится со стабом backend и проверяется на наличие ключевых
элементов (без пиксельных эталонов — только структура и поведение):
  A. режим установки: заголовок «Установка OVERTIMETAB», кнопка «Далее»,
     радиокнопки «Для меня» / «Для всех», путь по умолчанию;
  B. «Далее» открывает страницу места: галки «На рабочем столе» и
     «В меню Пуск», кнопка «Установить», поле пути;
  C. найденная копия: кнопка «Обновить там же» ведёт на прогресс и зовёт
     installExisting;
  D. пустой путь — ошибка, установки нет;
  E. режим удаления: заголовок «Удаление», галка данных выключена,
     кнопка «Удалить» зовёт uninstall(False);
  F. сигнал finishedOk переводит на «Готово».

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 python3 qa/test_setup_view.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "qa", "render")

from PySide6.QtCore import QObject, Property, Signal, Slot  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtQml import QQmlApplicationEngine  # noqa: E402


class StubBackend(QObject):
    """Стаб SetupBackend (Main.py): методы фиксируются вызовами.

    ВАЖНО (PySide6 6.11): в классе с Signal свойства Property(constant)
    ломают экспозицию объекта в QML («backend is null»). Как и в
    продовом Backend, все свойства здесь — notify-типа.
    """

    progressChanged = Signal()
    statusChanged = Signal()
    modeChanged = Signal()
    isDarkChanged = Signal()
    existingChanged = Signal()
    defaultDirChanged = Signal()
    finishedOk = Signal()
    finishedError = Signal(str)
    requestQuit = Signal()

    def __init__(self, mode="setup", existing=None):
        super().__init__()
        self._mode = mode
        self._existing = existing
        self.calls = []

    @Property(str, notify=modeChanged)
    def mode(self):
        return self._mode

    @Property(bool, notify=isDarkChanged)
    def isDarkTheme(self):
        return False

    @Property(bool, notify=existingChanged)
    def existingFound(self):
        return bool(self._existing)

    @Property(str, notify=existingChanged)
    def existingDir(self):
        return (self._existing or {}).get("dir", "")

    @Property(int, notify=existingChanged)
    def existingBuild(self):
        return int((self._existing or {}).get("build") or 0)

    @Property(bool, notify=existingChanged)
    def existingPerUser(self):
        return True

    @Property(str, notify=defaultDirChanged)
    def defaultDir(self):
        return "C:\\Users\\tester\\AppData\\Local\\Programs\\OVERTIMETAB"

    @Property(float, notify=progressChanged)
    def progress(self):
        return 0.0

    @Property(str, notify=statusChanged)
    def statusLine(self):
        return ""

    @Slot(bool, result=str)
    def dirFor(self, per_user):
        return ("C:\\Users\\tester\\AppData\\Local\\Programs\\OVERTIMETAB"
                if per_user else "C:\\Program Files\\OVERTIMETAB")

    @Slot(str, bool, bool, bool)
    def install(self, dest, per_user, desktop, start_menu):
        self.calls.append(("install", str(dest), bool(per_user),
                           bool(desktop), bool(start_menu)))

    @Slot()
    def installExisting(self):
        self.calls.append(("installExisting",))

    @Slot(bool)
    def uninstall(self, remove_data):
        self.calls.append(("uninstall", bool(remove_data)))

    @Slot(bool)
    def finish(self, run_after):
        self.calls.append(("finish", bool(run_after)))


def get_wizard(win):
    """Объект мастера (Item с property page) внутри окна-обёртки.

    Напрямую property("setup") PySide6 конвертировать не умеет (тип
    генерируется QML-движком), поэтому ищем ребёнка с нашим свойством.
    """
    from PySide6.QtCore import QObject as QO
    for c in [win] + win.findChildren(QO):
        v = c.property("page")
        if isinstance(v, str):
            return c
    raise AssertionError("не нашли объект мастера в окне")


def eff_visible(o) -> bool:
    """Эффективная видимость: своя property visible + все родители.

    QML-типы (AppButton_QMLTYPE_…) приходят из findChildren как QObject
    без QQuickItem-методов, поэтому читаем property("visible") — но
    property("visible") локальна, а страницы StackLayout прячут детей
    именно через родителя, потому идём по цепочке вверх.
    """
    cur = o
    while cur is not None:
        if cur.property("visible") is False:
            return False
        cur = cur.parent()
    return True


def find_text(root, text, visible_only=True):
    """Видимый элемент с property text/title (страницы StackLayout
    существуют всегда — учитываем видимость)."""
    def walk(o):
        from PySide6.QtCore import QObject as QO
        if o.property("text") == text or o.property("title") == text:
            if not visible_only or eff_visible(o):
                return o
        for c in o.findChildren(QO):
            r = walk(c)
            if r is not None:
                return r
        return None
    return walk(root)


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    results = {}

    def run(stub):
        engine = QQmlApplicationEngine()
        ctx = engine.rootContext()
        ctx.setContextProperty("backend", stub)
        # вход мастера — main_setup.qml из корня (как main.qml программы):
        # компонент из каталога-модуля components напрямую контекстных
        # свойств не видит
        qml = os.path.join(ROOT, "main_setup.qml")
        engine.load(qml)
        assert engine.rootObjects(), "мастер не загрузился"
        deadline = time.time() + 2.5
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.02)
        win = engine.rootObjects()[0]
        # держим ВСЕ движки и стабы: сборщик мусора не должен убивать
        # окна прошлых прогонов (bindings падают со «backend is null»-шумом)
        results.setdefault("engines", []).append(engine)
        results.setdefault("stubs", []).append(stub)
        return win

    # ── A. страница приветствия ──
    stub = StubBackend("setup")
    win = run(stub)
    w = get_wizard(win)
    assert find_text(win, "Далее") is not None, "нет кнопки «Далее»"
    assert stub.defaultDir and w.property("installDir") == stub.defaultDir
    assert not find_text(w, "Удалить"), "в установке нет кнопки удаления"
    print("A: приветствие установки, кнопка «Далее», путь по умолчанию ✓")

    # ── B. страница места установки ──
    btn = find_text(w, "Далее")
    btn.clicked.emit()
    for _ in range(20):
        app.processEvents()
        time.sleep(0.01)
    assert w.property("page") == "place", w.property("page")
    for label in ("На рабочем столе", "В меню «Пуск»"):
        assert find_text(win, label) is not None, "нет галки " + label
    assert find_text(win, "Установить") is not None, "нет кнопки «Установить»"
    assert find_text(win, "Для меня (рекомендуется) — без прав администратора")
    assert find_text(win, "Для всех пользователей (папка Program Files)")
    print("B: место установки — режимы, галки ярлыков, кнопка «Установить» ✓")

    # ── D. пустой путь — ошибка, установки нет ──
    # поле пути: QQuickTextInput не экспортирован в PySide6.QtQuick —
    # ищем по имени класса
    from PySide6.QtCore import QObject as QO2
    tins = [c for c in w.findChildren(QO2)
            if c.metaObject().className() == "QQuickTextInput"]
    assert tins, "нет поля пути"
    tins[0].setProperty("text", "   ")
    btn2 = find_text(w, "Установить")
    btn2.clicked.emit()
    for _ in range(10):
        app.processEvents()
    assert w.property("errorText") and not stub.calls, (w.property("errorText"), stub.calls)
    assert w.property("page") == "place"
    print("D: пустой путь — ошибка на месте, установка не стартует ✓")

    # ── C. найденная копия → «Обновить там же» ──
    stub2 = StubBackend("setup", existing={
        "dir": "C:\\Users\\tester\\AppData\\Local\\Programs\\OVERTIMETAB",
        "build": 242, "version": "2.0.0-ALPHA.242"})
    win2 = run(stub2)
    w2 = get_wizard(win2)
    up = find_text(w2, "Обновить там же")
    assert up is not None, "нет кнопки «Обновить там же»"
    assert find_text(w2, "Выбрать другое место") is not None
    up.clicked.emit()
    for _ in range(20):
        app.processEvents()
        time.sleep(0.01)
    assert w2.property("page") == "progress", w2.property("page")
    assert stub2.calls == [("installExisting",)], stub2.calls
    print("C: найденная копия — «Обновить там же» ведёт на прогресс ✓")

    # сигнал finishedOk → «Готово»
    stub2.finishedOk.emit()
    for _ in range(20):
        app.processEvents()
        time.sleep(0.01)
    assert w2.property("page") == "done", w2.property("page")
    assert find_text(w2, "Запустить программу") is not None
    print("F: finishedOk → страница «Готово» с галкой запуска ✓")

    # ── E. режим удаления ──
    stub3 = StubBackend("uninstall")
    win3 = run(stub3)
    assert win3.title() == "Удаление OVERTIMETAB", win3.title()
    w3 = get_wizard(win3)
    assert w3.property("page") == "confirmRemove", w3.property("page")
    gal = find_text(w3, "Удалить также данные сотрудников — «Документы\\OverTimeTab»")
    assert gal is not None, "нет галки удаления данных"
    assert gal.property("checked") is False, "галка данных должна быть выключена"
    btn3 = find_text(w3, "Удалить")
    assert btn3 is not None
    btn3.clicked.emit()
    for _ in range(20):
        app.processEvents()
        time.sleep(0.01)
    assert w3.property("page") == "progress", w3.property("page")
    assert stub3.calls == [("uninstall", False)], stub3.calls
    print("E: удаление — подтверждение, данные не трогаем по умолчанию ✓")

    print("═══ МАСТЕР УСТАНОВКИ: ОКНА И ПОВЕДЕНИЕ ЦЕЛЫ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
