#!/usr/bin/env python3
"""Тест живой виньетки окна «Что нового».

Сверху окна после обновления появляется маленькая живая сцена,
показывающая «показывающие» новинки выпуска. Питон отдаёт список сцен
(backend.whatsNewScenes); если их несколько, виньетка листает их сама:
одна проигралась — наступает следующая. Точки под сценой показывают,
какая сейчас, и позволяют выбрать любую.

Проверяется:
  A. Виньетка показывается, только когда есть что рассказать:
     записи есть — сцена видна; «всё уже видели» — сцены нет;
  H. Окно подстраивается по высоте под виньетку и записи (не микро-
     окошко), и дорастает при смене содержимого;
  B. Сцена «волна»: семь мини-дней, сюжет идёт по кругу — метки
     собираются по одной и растворяются; курсор гаснет после прохода;
     единственная сцена не листается сама;
  C. Рендер: мини-дни и собранные метки видны на экране (пиксели);
  K. Карусель: сцены сменяют друг друга сами по очереди; точки
     показывают текущую и выбирают любую вручную; возврат к сцене
     перезапускает её сюжет;
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

        # появление виньетки: проявление и подъём
        assert vign.property("opacity") > 0.9, \
            "виньетка не проявилась после открытия окна"
        print("A: окно открыто — виньетка сверху, проявилась ✓")

        # ── H. Высота окна подстраивается под содержимое ──
        # (окно 700x600: потолок 500; виньетка + запись ≈ 460–500)
        h1 = dlg.property("height")
        assert h1 > 420, f"окно микро (высота {h1:.0f}), а не по контенту"
        # записей стало больше — окно дорастает до потолка
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
        # вернём одну запись для сюжетных проверок
        dlg.close()
        settle(0.3)
        stub._wn = make_entries(1)
        stub.whatsNewChanged.emit()
        settle(0.2)
        dlg.open()
        settle(0.4)
        print(f"H: высота окна по контенту ({h1:.0f} → {h2:.0f} по записям) ✓")

        # ── B. Сцена «волна»: сюжет идёт по кругу ──
        wave = find_by_name(dlg, "vignetteWaveScene")
        assert wave is not None and wave.property("visible") is True, \
            "сцена-волна не показана"
        cells = scene_items(wave, "vignetteMiniCell")
        badges = scene_items(wave, "vignetteMiniBadge")
        assert len(cells) == 7, f"мини-дней {len(cells)}, а не 7"
        assert len(badges) == 7, f"меток {len(badges)}, а не 7"

        # единственная сцена — точек не видно
        dots_row = find_by_name(dlg, "vignetteDots")
        assert dots_row is not None and dots_row.property("visible") is False, \
            "точки у единственной сцены не нужны"

        # метки собираются по одной: к 2,4 с все на месте (волна кончилась
        # в 1,9 с), к 4,3 с все растворились (пауза сюжета до 5,4 с)
        t0 = time.time()
        settle(2.4)
        badges = scene_items(wave, "vignetteMiniBadge")
        opacities = [b.property("opacity") for b in badges]
        assert all(op > 0.99 for op in opacities), \
            f"метки волны не собрались: {opacities}"
        cursor = scene_items(wave, "vignetteCursor")[0]
        assert cursor.property("opacity") < 0.1, \
            "курсор не погас после прохода по дням"
        print("B: волна — метки собрались по одной, курсор погас ✓")

        # ── C. Рендер: мини-дни и метки видны на экране ──
        img = qwin.grabWindow()
        assert img.width() > 0 and img.height() > 0, "пустой снимок окна"

        from PySide6.QtCore import QPointF

        def center_color(item):
            pt = item.mapToScene(QPointF(item.width() / 2, item.height() / 2))
            return img.pixelColor(int(pt.x()), int(pt.y()))

        # выходной мини-день (числа 10–11) — красноватая подложка,
        # красный канал выше синего
        px_weekend = center_color(cells[5])
        assert px_weekend.red() > px_weekend.blue() + 8, \
            f"выходной мини-день не подсвечен: {px_weekend.name()}"
        # метка «О» — оранжевая буква в центре
        px_badge = center_color(badges[3])
        assert px_badge.red() > px_badge.blue() + 8, \
            f"метка «О» не видна на экране: {px_badge.name()}"
        print(f"C: рендер — выходной день {px_weekend.name()}, "
              f"метка «О» {px_badge.name()} ✓")

        # сюжет по кругу: метки растворились к паузе; сцена одна —
        # авто-листания нет, волна остаётся
        settle(4.3 - (time.time() - t0))
        badges = scene_items(wave, "vignetteMiniBadge")
        opacities = [b.property("opacity") for b in badges]
        assert all(op < 0.1 for op in opacities), \
            f"метки не растворились к паузе сюжета: {opacities}"
        assert vign.property("currentScene") == "wave", \
            "единственная сцена не должна листаться сама"
        print("B2: сюжет идёт по кругу, единственная сцена не листается ✓")

        # ── K. Карусель: сцены сменяют друг друга сами ──
        # Короткая пауза — НАМЕРЕННО: повторное открытие раньше, чем
        # истечёт страховка закрытия (600 мс), когда-то убивало окно
        # посреди входного перехода
        dlg.close()
        settle(0.3)
        stub._scenes = ["wave", "neutral"]
        stub.whatsNewChanged.emit()
        settle(0.2)
        dlg.open()
        t0 = time.time()
        settle(0.5)                     # входной переход окна (250 мс)
        assert vign.property("active") is True, "окно не открылось"
        assert vign.property("currentScene") == "wave", \
            "после открытия — первая сцена"
        dots = scene_items(vign, "vignetteDot")
        assert len(dots) == 2, f"точек {len(dots)}, а не 2"
        assert dots[0].property("visible") is True and \
            dots[1].property("visible") is True, "точки не видны"

        settle(3.0)                     # ≈3,5 с от открытия: волна ещё идёт
        assert vign.property("currentScene") == "wave", \
            "волна не проигралась ещё полностью"
        settle(1.2)                     # ≈4,7 с: авто-листание настало
        assert vign.property("currentScene") == "neutral", \
            "после волны должна наступить нейтральная сцена сама"
        print("K: сцены листаются сами — после волны наступила "
              "нейтральная ✓")

        # ── D. Нейтральная сцена: маскот и конфетти ──
        neutral = find_by_name(dlg, "vignetteNeutralScene")
        assert neutral is not None and neutral.property("visible") is True, \
            "нейтральная сцена не показана"
        mascot = vign.neutralMascot()
        assert mascot is not None, "маскот не найден в нейтральной сцене"
        assert mascot.property("visible") is True
        pieces = scene_items(neutral, "vignetteConfetti")
        assert len(pieces) == 8, f"конфетти {len(pieces)}, а не 8"

        # конфетти медленно сыплется (первый лоскут стартовал и виден)
        settle(1.4)
        top = max(p.property("opacity") for p in pieces)
        assert top > 0.05, "конфетти не сыплется"

        # рендер: первый лоскут — синий цвет программы
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
        badges = scene_items(wave, "vignetteMiniBadge")
        fresh = [b.property("opacity") for b in badges]
        assert all(op < 0.1 for op in fresh), \
            f"возврат к сцене не перезапустил сюжет: {fresh}"
        # активная точка — шире остальных (растёт через Behavior, ждём)
        settle(0.25)
        assert dots[0].property("width") > dots[1].property("width"), \
            "активная точка не выделена"
        # авто-листание продолжается с выбранной сцены
        settle(0.4)
        assert vign.property("currentScene") == "wave"
        settle(3.8)
        assert vign.property("currentScene") == "neutral", \
            "авто-листание не продолжилось после ручного выбора"
        print("K2: точка выбирает сцену, сюжет перезапускается, "
              "авто-листание продолжается ✓")

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

        print("═══ ЖИВАЯ ВИНЬЕТКА «ЧТО НОВОГО»: КАРУСЕЛЬ СЦЕН ═══")
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
