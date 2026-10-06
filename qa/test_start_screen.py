#!/usr/bin/env python3
"""Тест экрана выбора базы и анимации заставки (главный main.qml).

Сценарий юзера: клик по базе раскрывает панель-заставку на весь экран,
при успехе она улетает вправо с инерцией, открывая интерфейс открытой
базы; при неудаче (файла нет) возвращается на место — программа не виснет.

Инварианты (брак сборки 279: панель нулевой высоты из-за сломанного
якоря «Cannot anchor to an item that isn't a parent or sibling»
и второй маскот под улетающей панелью):
  - панель-заставка всегда полноценной высоты;
  - в каждый момент виден ровно ОДИН маскот;
  - при успехе страница меняется мгновенно — под улетающей панелью
    готовый интерфейс базы.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen python3 qa/test_start_screen.py
"""
import os
import sys
import time
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Signal, Property, Slot  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtQml import QQmlApplicationEngine  # noqa: E402
from PySide6.QtQuick import QQuickItem, QQuickWindow  # noqa: E402


class StubBackend(QObject):
    databaseOpened = Signal()
    databaseOpenFailed = Signal(str)
    itemDeleted = Signal(str, str)
    showToast = Signal(str, str)
    dbListChanged = Signal()
    activeDepartmentNameChanged = Signal()
    trayHintChanged = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._fail = False
        self._tray = False

    @Property(list, notify=dbListChanged)
    def dbList(self):
        return [{"name": "Тестовое подразделение", "path": "C:/базы/test.sqlite"}]

    @Property(str, notify=activeDepartmentNameChanged)
    def activeDepartmentName(self):
        return "" if self._fail else "Тестовое подразделение"

    @Property(str, constant=True)
    def currentPeriodText(self):
        return "Октябрь 2026"

    @Property(bool, constant=True)
    def reminderEnabled(self):
        return True

    @Property(bool, constant=True)
    def startHidden(self):
        return False

    @Property(bool, notify=trayHintChanged)
    def trayHintWasShown(self):
        return self._tray

    @Property(list, constant=True)
    def hotkeysList(self):
        return []

    @Slot()
    def setTrayHintShown(self):
        self._tray = True
        self.trayHintChanged.emit()

    @Slot(str)
    def openDatabase(self, path):
        if self._fail:
            self.showToast.emit("Не удалось открыть базу: файл не найден", "error")
            self.databaseOpenFailed.emit("файл не найден")
        else:
            self.databaseOpened.emit()

    @Slot(str)
    def attachDatabase(self, url):
        pass

    @Slot(str)
    def changeDbDirectory(self, url):
        pass

    @Slot(str)
    def openDbFolder(self, path):
        pass

    @Slot()
    def exportToExcel(self):
        pass

    @Slot()
    def quickPrint(self):
        pass

    @Slot()
    def undoAction(self):
        pass

    @Slot()
    def redoAction(self):
        pass

    @Slot()
    def scanForUpdates(self):
        pass


def walk_items(root, out=None):
    """Все QQuickItem ниже root (включая contentItem окна)."""
    if out is None:
        out = []
    try:
        kids = root.contentItem().childItems()
    except Exception:
        try:
            kids = root.childItems()
        except Exception:
            kids = []
    for child in kids:
        out.append(child)
        walk_items(child, out)
    return out


def qquick_window(app):
    for w in app.allWindows():
        if isinstance(w, QQuickWindow):
            return w
    raise AssertionError("QQuickWindow не найден")


def visible_mascots(win):
    out = []
    for i in walk_items(win):
        try:
            cn = i.metaObject().className() or ""
        except Exception:
            continue
        if ("Mascot" in cn or "EmptyMascot" in cn) and i.isVisible():
            out.append(i)
    return out


def find_cover(win):
    for i in walk_items(win):
        if i.property("path") is not None and hasattr(i, "width"):
            return i
    return None


def pump(app, seconds, win, collect):
    """Съёмка кадров: шторку держим по ссылке (обход дерева дорог и
    размазывает тайминги), маскотов считаем каждый третий кадр."""
    deadline = time.time() + seconds
    cover = None
    i = 0
    while time.time() < deadline:
        app.processEvents()
        time.sleep(0.012)
        if cover is None or not cover.isVisible():
            cover = find_cover(win)
        if cover is not None and cover.isVisible():
            masc = len(visible_mascots(win)) if i % 3 == 0 else -1
            collect.append((cover.x(), cover.width(), cover.height(), True, masc))
        else:
            collect.append((0.0, 0.0, 0.0, False, 0))
        i += 1


def load(app, fail):
    stub = StubBackend()
    stub._fail = fail
    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("backend", stub)
    engine.load(str(ROOT / "main.qml"))
    assert engine.rootObjects(), "main.qml не загрузился"
    for _ in range(6):
        app.processEvents()
    win = engine.rootObjects()[0]
    if not isinstance(win, QQuickWindow):
        win = qquick_window(app)
    page = None
    for i in walk_items(win):
        if hasattr(i, "openDbAnimated"):
            page = i
            break
    assert page is not None, "страница выбора базы не найдена"
    return stub, engine, win, page


def main() -> int:
    app = QGuiApplication.instance() or QGuiApplication(sys.argv)

    # ── A. успех: рост на весь экран → улёт вправо с инерцией ──
    stub, engine, win, page = load(app, fail=False)
    page.openDbAnimated("C:/базы/test.sqlite")
    frames = []
    pump(app, 2.2, win, frames)
    assert frames, "кадров нет"
    for x, w, h, vis, masc in frames:
        if vis:
            assert h > 400, "панель-заставка потеряла высоту: %s" % ((x, w, h),)
        if masc >= 0:
            assert masc <= 1, "видно несколько маскотов: %s" % masc
    grew = max(f[1] for f in frames if f[3])
    assert grew > win.width() * 0.9, "панель не раскрылась на весь экран: %s" % grew
    flew = max(f[0] for f in frames if f[3])
    assert flew > win.width() * 0.8, "панель не улетела вправо: %s" % flew
    # «живая» инерция — бросок как в дизайне: замах назад и кривая с
    # выходом на полной скорости (проверяем по коду: замах в 90 мс
    # покадровая съёмка стенда ловит нестабильно)
    qml_src = (ROOT / "main.qml").read_text(encoding="utf-8")
    assert "to: -16; duration: 100" in qml_src, "пропал замах назад"
    # кривая броска — OutQuint 1150 мс: взрывной старт и планирование
    # (~400 мс видимого полёта). Easing.Bezier в этой версии Qt молча
    # падает в Linear (замерено прототипом — обе «кривые» 281 выходили
    # линейными) — запрещаем навсегда
    assert "easing.type: Easing.OutQuint" in qml_src, "пропал бросок OutQuint"
    assert "duration: 1150" in qml_src, "пропала длительность броска"
    assert "bezierCurve" not in qml_src, "Bezier молча линейный — запрещён"
    assert "contentOpacity" in qml_src, "пропало гаснущее содержимое шторки"
    # бросок стартует ПОСЛЕ паузы-бит: тяжёлое построение рабочего экрана
    # иначе замораживает анимацию на полпути (брак 281 — панель исчезала)
    assert "flightDelay" in qml_src, "пропала пауза на построение экрана"
    # динамика: замах назад и быстрый улёт (кадров «в полёте» немного,
    # крадущийся хвост у края — признак тормозящей кривой)
    xs = [f[0] for f in frames if f[3]]
    assert any(x < -4 for x in xs), "нет замаха назад: %s" % xs[:8]
    flying = [x for x in xs if 0 < x < win.width()]
    assert 6 <= len(flying) <= 48, "кадров полёта вне диапазона: %s" % flying
    # страница сменилась мгновенно: под улетающей панелью интерфейс базы
    workspace = None
    for i in walk_items(win):
        try:
            if i.property("calendarPanel") is not None:
                workspace = i
                break
        except RuntimeError:
            # свойство есть, но PySide6 не конвертирует тип CalendarWorkspace
            workspace = i
            break
    if workspace is None:
        classes = []
        for i in walk_items(win):
            try:
                classes.append(i.metaObject().className() or "?")
            except Exception:
                pass
        raise AssertionError("под панелью не рабочий интерфейс базы; предметы: %s"
                             % classes[:40])
    # интерфейс базы «подъехал»: проявился и поднялся на место
    assert abs(float(workspace.property("opacity")) - 1.0) < 0.01, (
        "интерфейс не проявился: %s" % workspace.property("opacity"))
    assert abs(float(workspace.property("y"))) < 0.01, (
        "интерфейс не поднялся на место: %s" % workspace.property("y"))
    assert not frames[-1][3], "панель не скрылась после улёта"
    print("A: рост на весь экран → улёт вправо; маскот всегда один; "
          "под панелью готовый интерфейс базы ✓")
    engine.deleteLater()

    # ── B. неудача (авто-открытие единственной пропавшей базы) ──
    stub2, engine2, win2, page2 = load(app, fail=True)
    # страница сама пытается открыть единственную базу через ~100 мс
    frames2 = []
    pump(app, 2.5, win2, frames2)
    assert frames2, "кадров неудачи нет"
    for x, w, h, vis, masc in frames2:
        if vis:
            assert h > 400, "панель потеряла высоту при возврате: %s" % ((x, w, h),)
        if masc >= 0:
            assert masc <= 1, "несколько маскотов при возврате: %s" % masc
    grew2 = max((f[1] for f in frames2 if f[3]), default=0)
    assert grew2 > win2.width() * 0.9, "панель не раскрылась: %s" % grew2
    assert not frames2[-1][3], "панель не скрылась после возврата"
    assert len(visible_mascots(win2)) == 1, "экран выбора не восстановился"
    assert win2.property("startTransition") is False, "переход не разблокирован"
    # повторный клик возможен
    page2.openDbAnimated("C:/базы/test.sqlite")
    frames3 = []
    pump(app, 1.5, win2, frames3)
    assert any(f[3] for f in frames3), "повторный клик не запустил заставку"
    print("B: неудача — панель вернулась на место, экран выбора жив, "
          "повторный клик работает ✓")

    print("═══ ЭКРАН ВЫБОРА БАЗЫ: ЗАСТАВКА РАСТЁТ, ЛЕТИТ, ВОЗВРАЩАЕТСЯ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
