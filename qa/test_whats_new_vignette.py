#!/usr/bin/env python3
"""Тест живой виньетки окна «Что нового» (сцена-театр).

Сверху окна после обновления появляется маленькая живая сцена. Спектакль
собран из актёров (дни, курсор, меню статусов, мини-окно) и «плеера
тактов»: сцена — это список «что за чем происходит», плеер проигрывает
его по кругу. Сцены листаются сами, точки под сценой выбирают любую.

Проверяется:
  A. Виньетка показывается, только когда есть что рассказать;
  H. Окно подстраивается по высоте под виньетку и записи;
  B. Сцена «волна»: ВЕСЬ сюжет — курсор ведёт по дням, открывается
     меню статусов, «О» прожимается, меню складывается, метки
     собираются волной; сюжет по кругу; единственная сцена
     не листается сама;
  C. Рендер: мини-дни и собранные метки видны на экране (пиксели);
  K. Карусель: сцены сменяют друг друга сами; точки выбирают любую;
     возврат к сцене перезапускает сюжет; авто-листание продолжается;
  W. Актёр «мини-окно»: открывается, кнопка прожимается, закрывается
     (грамматика для будущих сцен);
  D. Нейтральная сцена: маскот и конфетти;
  E. Закрытие окна зовёт ackWhatsNew (порция считается прочитанной).

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 python3 qa/test_whats_new_vignette.py
"""
import os
import sys
import time
import types

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Property, Signal, Slot  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtQml import QQmlApplicationEngine  # noqa: E402
from PySide6.QtQuick import QQuickWindow  # noqa: E402


def make_entries(n):
    out = []
    for i in range(n):
        out.append({
            "hasAdded": True, "hasChanged": i % 2 == 0,
            "hasFixed": False, "hasRemoved": False,
            "addedText": "<br>".join("• Запись %d, длинная enough для "
                                     "переноса на вторую строчку" % j
                                     for j in range(2)),
            "changedText": "• Изменение %d" % i if i % 2 == 0 else "",
            "fixedText": "", "removedText": "",
            "version": "", "date": "",
        })
    return out


class StubBackend(QObject):
    """Стаб бэкенда для окна «Что нового»."""

    whatsNewChanged = Signal()

    def __init__(self):
        super().__init__()
        self._wn = make_entries(1)
        self._scenes = ["wave"]
        self.acked = False

    @Property(list, notify=whatsNewChanged)
    def whatsNew(self):
        return self._wn

    @Property(list, notify=whatsNewChanged)
    def whatsNewScenes(self):
        return self._scenes

    @Property(bool, notify=whatsNewChanged)
    def whatsNewFresh(self):
        return True

    @Slot()
    def ackWhatsNew(self):
        self.acked = True


WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI

ApplicationWindow {
    id: w
    width: 700; height: 600; visible: true
    AppUI.WhatsNewDialog { id: dlg; objectName: "whatsNewDlg" }
}
"""


def find_by_name(root, name):
    return root.findChild(QObject, name)


def iter_items(item):
    """Обход визуального дерева: делегаты Repeater сидят в childItems,
    а не в QObject-цепочке — обычный findChild их не находит.
    Внутрь маскота не спускаемся: его эффекты роняют конвертеры PySide."""
    for c in item.childItems():
        yield c
        if c.objectName() == "vignetteMascot":
            continue
        yield from iter_items(c)


def scene_items(root_item, name):
    return [c for c in iter_items(root_item) if c.objectName() == name]


def wait_until(app, cond, timeout=3.0):
    """Крутит события, пока не выполнится условие (или не выйдет время)."""
    t = time.time() + timeout
    while time.time() < t:
        app.processEvents()
        time.sleep(0.02)
        if cond():
            return True
    return cond()


def main() -> int:
    qml_path = os.path.join(ROOT, "_render_vignette.qml")
    with open(qml_path, "w", encoding="utf-8") as f:
        f.write(WRAPPER)

    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()          # ссылка обязательна (см. test_day_select)
    stub = StubBackend()
    ctx.setContextProperty("backend", stub)

    def settle(seconds):
        t = time.time() + seconds
        while time.time() < t:
            app.processEvents()
            time.sleep(0.01)

    try:
        engine.load(qml_path)
        assert engine.rootObjects(), "окно не загрузилось"
        settle(0.4)
        win = engine.rootObjects()[0]
        dlg = find_by_name(win, "whatsNewDlg")
        assert dlg is not None, "окно «Что нового» не найдено"
        qwin = next(x for x in app.topLevelWindows()
                    if isinstance(x, QQuickWindow))

        dlg.open()
        settle(0.6)
        vign = find_by_name(dlg, "whatsNewVignette")
        assert vign is not None, "виньетка не найдена"
        assert vign.property("active") is True, "виньетка не активна"
        assert vign.property("visible") is True, "виньетка не видна"
        assert vign.property("currentScene") == "wave"
        assert vign.property("opacity") > 0.9, \
            "виньетка не проявилась после открытия окна"
        stage = find_by_name(dlg, "whatsNewStage")
        assert stage is not None, "сцена-театр не найдена"
        assert stage.property("playing") is True, "плеер не играет"
        print("A: окно открыто — виньетка с театром, плеер играет ✓")

        # ── H. Высота окна подстраивается под содержимое ──
        h1 = dlg.property("height")
        assert h1 > 420, f"окно микро (высота {h1:.0f}), а не по контенту"
        dlg.close()
        settle(0.3)
        stub._wn = make_entries(5)
        stub.whatsNewChanged.emit()
        settle(0.2)
        dlg.open()
        settle(0.6)
        h2 = dlg.property("height")
        assert h2 >= h1 and h2 >= 495, \
            f"окно не доросло до потолка (было {h1:.0f}, стало {h2:.0f})"
        dlg.close()
        settle(0.3)
        stub._wn = make_entries(1)
        stub.whatsNewChanged.emit()
        settle(0.2)
        dlg.open()
        settle(0.4)
        print(f"H: высота окна по контенту ({h1:.0f} → {h2:.0f} по записям) ✓")

        # ── B. Сцена «волна»: весь сюжет, с меню статусов ──
        stage = find_by_name(dlg, "whatsNewStage")
        cells = scene_items(stage, "vignetteMiniCell")
        badges = scene_items(stage, "vignetteMiniBadge")
        assert len(cells) == 7, f"мини-дней {len(cells)}, а не 7"
        assert len(badges) == 7, f"меток {len(badges)}, а не 7"
        menu = find_by_name(dlg, "vignetteMenu")
        assert menu is not None, "меню статусов не найдено"
        cursor = scene_items(stage, "vignetteCursor")[0]

        # единственная сцена — точек не видно
        assert find_by_name(dlg, "vignetteDots").property("visible") is False

        # сценарий длинный: полный прогон больше 4 с (карусель по нему)
        total = stage.property("totalMs")
        assert total > 4000, f"прогон сцены подозрительно короткий: {total}"
        assert vign.property("sceneCycleMs") >= total + 400

        # такт 3–6: меню открыто, курсор у него
        settle(1.2)                     # ≈1,6 с от открытия (1,3 с сцены)
        assert menu.property("opacity") > 0.9, \
            "меню статусов не открылось после выделения дней"

        # такты 7–8: меню сложилось, метки собираются волной
        settle(1.8)                     # ≈3,4 с открытия: волна собралась (3,1 сцены)
        badges = scene_items(stage, "vignetteMiniBadge")
        opacities = [b.property("opacity") for b in badges]
        assert all(op > 0.99 for op in opacities), \
            f"метки волны не собрались: {opacities}"
        assert menu.property("opacity") < 0.05, \
            "меню не сложилось после выбора статуса"
        assert cursor.property("opacity") < 0.1, \
            "курсор не погас после выбора"
        print("B: сюжет — выделение, меню, «О», волна меток ✓")

        # ── C. Рендер: мини-дни и метки видны на экране ──
        img = qwin.grabWindow()
        assert img.width() > 0 and img.height() > 0, "пустой снимок окна"

        from PySide6.QtCore import QPointF

        def center_color(item):
            pt = item.mapToScene(QPointF(item.width() / 2, item.height() / 2))
            return img.pixelColor(int(pt.x()), int(pt.y()))

        px_weekend = center_color(cells[5])
        assert px_weekend.red() > px_weekend.blue() + 8, \
            f"выходной мини-день не подсвечен: {px_weekend.name()}"
        px_badge = center_color(badges[3])
        assert px_badge.red() > px_badge.blue() + 8, \
            f"метка «О» не видна на экране: {px_badge.name()}"
        print(f"C: рендер — выходной день {px_weekend.name()}, "
              f"метка «О» {px_badge.name()} ✓")

        # такт 10: метки растворились, выделение снято (ждём конца)
        badges = scene_items(stage, "vignetteMiniBadge")
        assert wait_until(app,
            lambda: all(b.property("opacity") < 0.1 for b in badges), 2.5), \
            "метки не растворились к паузе сюжета: " \
            f"{[b.property('opacity') for b in badges]}"
        assert vign.property("currentScene") == "wave", \
            "единственная сцена не должна листаться сама"

        # сюжет по кругу: новый прогон уже начался (курсор вернётся)
        assert wait_until(app, lambda: cursor.property("opacity") > 0.5, 3.0), \
            "сюжет не пошёл по кругу — курсор не появился снова"
        assert vign.property("currentScene") == "wave"
        print("B2: сюжет по кругу, единственная сцена не листается ✓")

        # ── K. Карусель: сцены сменяют друг друга сами ──
        dlg.close()
        settle(0.3)                     # намеренно коротко: гонка страховки
        stub._scenes = ["wave", "neutral"]
        stub.whatsNewChanged.emit()
        settle(0.2)
        dlg.open()
        settle(0.5)
        assert vign.property("currentScene") == "wave", \
            "после открытия — первая сцена"
        dots = scene_items(vign, "vignetteDot")
        assert len(dots) == 2 and dots[0].property("visible") is True, \
            "точек нет, а сцен две"

        settle(4.0)                     # ≈4,5 с: волна ещё играет
        assert vign.property("currentScene") == "wave", \
            "волна ещё не проигралась, а сцена уже сменилась"
        assert wait_until(app,
            lambda: vign.property("currentScene") == "neutral", 3.0), \
            "после волны должна наступить нейтральная сцена сама"
        print("K: сцены листаются сами — после волны наступила "
              "нейтральная ✓")

        # ── D. Нейтральная сцена: маскот и конфетти ──
        neutral = find_by_name(dlg, "vignetteNeutralScene")
        assert neutral is not None, "нейтральная сцена не найдена"
        assert wait_until(app, lambda: neutral.property("visible") is True, 1.5), \
            "нейтральная сцена не показана"
        mascot = vign.neutralMascot()
        assert mascot is not None, "маскот не найден в нейтральной сцене"
        assert mascot.property("visible") is True
        pieces = scene_items(neutral, "vignetteConfetti")
        assert len(pieces) == 8, f"конфетти {len(pieces)}, а не 8"

        settle(1.4)
        top = max(p.property("opacity") for p in pieces)
        assert top > 0.05, "конфетти не сыплется"
        piece0 = min(pieces, key=lambda p: p.property("baseX"))
        img = qwin.grabWindow()
        pt = piece0.mapToScene(QPointF(piece0.width() / 2,
                                       piece0.height() / 2))
        px_piece = img.pixelColor(int(pt.x()), int(pt.y()))
        assert px_piece.blue() > px_piece.red() + 8, \
            f"конфетти не видно на экране: {px_piece.name()}"
        print(f"D: нейтральная сцена — маскот, конфетти сыплется "
              f"(пиксель {px_piece.name()}) ✓")

        # ── K2. Точка выбирает сцену вручную; сюжет перезапускается ──
        vign.goToScene(0)
        app.processEvents()
        assert vign.property("currentScene") == "wave", \
            "точка не вернула сцену-волну"
        badges = scene_items(stage, "vignetteMiniBadge")
        fresh = [b.property("opacity") for b in badges]
        assert all(op < 0.1 for op in fresh), \
            f"возврат к сцене не перезапустил сюжет: {fresh}"
        settle(0.25)
        assert dots[0].property("width") > dots[1].property("width"), \
            "активная точка не выделена"
        assert wait_until(app,
            lambda: vign.property("currentScene") == "neutral", 8.0), \
            "авто-листание не продолжилось после ручного выбора"
        print("K2: точка выбирает сцену, сюжет перезапускается, "
              "авто-листание продолжается ✓")

        # ── W. Актёр «мини-окно»: грамматика для будущих сцен ──
        stage.stopPlay()
        stage.windowOpen()
        settle(0.15)
        mini_window = find_by_name(dlg, "vignetteWindow")
        assert mini_window is not None, "мини-окно не найдено"
        assert mini_window.property("opacity") > 0.9 and \
            mini_window.property("scale") > 0.9, "мини-окно не открылось"
        stage.windowPressSave()
        settle(0.35)
        save_btn = find_by_name(dlg, "vignetteSaveButton")
        assert save_btn is not None and \
            abs(save_btn.property("pressScale") - 1.0) < 0.05, \
            "прожатие кнопки не отработало"
        stage.windowClose()
        settle(0.15)
        assert mini_window.property("opacity") < 0.05, \
            "мини-окно не закрылось"
        print("W: мини-окно — открытие, прожатие, закрытие ✓")

        # ── A2. Записей нет («всё уже видели») — сцены нет ──
        stub._wn = []
        stub.whatsNewChanged.emit()
        settle(0.3)
        assert vign.property("visible") is False, \
            "без записей виньетка должна спрятаться"
        print("A2: «всё уже видели» — виньетка спрятана ✓")

        # ── E. Закрытие окна — порция прочитана ──
        dlg.close()
        settle(0.3)
        assert stub.acked, "закрытие окна не отметило порцию прочитанной"
        print("E: закрытие окна отметило порцию прочитанной ✓")

        print("═══ ЖИВАЯ ВИНЬЕТКА «ЧТО НОВОГО»: ТЕАТР СЦЕНАРИЕВ ═══")
    finally:
        try:
            os.remove(qml_path)
        except OSError:
            pass

    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    # PySide6 6.12 падает при финализации интерпретатора после QML-сцены
    # (none_dealloc на выходе) — выходим сразу после проверок
    os._exit(rc)
