#!/usr/bin/env python3
"""Тест живой виньетки окна «Что нового».

Сверху окна после обновления появляется маленькая живая сцена,
показывающая главное изменение выпуска. Сцену выбирает питон
по сборке (backend.whatsNewScene): «wave» — курсор проводит по
семи дням, метки отпуска собираются по одной; «neutral» — маскот
и лёгкое конфетти в цветах программы.

Проверяется:
  A. Виньетка показывается, только когда есть что рассказать:
     записи есть — сцена видна; «всё уже видели» — сцены нет;
  B. Сцена «волна»: семь мини-дней, сюжет идёт по кругу — метки
     собираются по одной и растворяются; курсор гаснет после
     прохода;
  C. Рендер: мини-дни и собранные метки видны на экране (пиксели);
  D. Нейтральная сцена: маскот и конфетти, сцены сменяют друг
     друга по свойству backend.whatsNewScene;
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


class StubBackend(QObject):
    """Стаб бэкенда для окна «Что нового»."""

    whatsNewChanged = Signal()

    def __init__(self):
        super().__init__()
        self._wn = [{
            "hasAdded": True, "hasChanged": False,
            "hasFixed": False, "hasRemoved": False,
            "addedText": "• Проверочная запись", "changedText": "",
            "fixedText": "", "removedText": "",
            "version": "", "date": "",
        }]
        self._scene = "wave"
        self.acked = False

    @Property(list, notify=whatsNewChanged)
    def whatsNew(self):
        return self._wn

    @Property(str, notify=whatsNewChanged)
    def whatsNewScene(self):
        return self._scene

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

        dlg.open()
        settle(0.4)
        vign = find_by_name(dlg, "whatsNewVignette")
        assert vign is not None, "виньетка не найдена"
        assert vign.property("active") is True, "виньетка не активна"
        assert vign.property("visible") is True, "виньетка не видна"

        # появление виньетки: проявление и подъём
        assert vign.property("opacity") > 0.9, \
            "виньетка не проявилась после открытия окна"
        print("A: окно открыто — виньетка сверху, проявилась ✓")

        # ── B. Сцена «волна»: сюжет идёт по кругу ──
        wave = find_by_name(dlg, "vignetteWaveScene")
        assert wave is not None and wave.property("visible") is True, \
            "сцена-волна не показана"
        cells = scene_items(wave, "vignetteMiniCell")
        badges = scene_items(wave, "vignetteMiniBadge")
        assert len(cells) == 7, f"мини-дней {len(cells)}, а не 7"
        assert len(badges) == 7, f"меток {len(badges)}, а не 7"

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
        qwin = next(x for x in app.topLevelWindows()
                    if isinstance(x, QQuickWindow))
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

        # сюжет по кругу: метки растворились к паузе
        settle(4.3 - (time.time() - t0))
        opacities = [b.property("opacity") for b in badges]
        assert all(op < 0.1 for op in opacities), \
            f"метки не растворились к паузе сюжета: {opacities}"
        print("B2: сюжет идёт по кругу — метки растворились и соберутся "
              "заново ✓")

        # ── D. Нейтральная сцена: маскот и конфетти ──
        stub._scene = "neutral"
        stub.whatsNewChanged.emit()
        settle(0.4)
        assert wave.property("visible") is False, \
            "сцена-волна не уступила место нейтральной"
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

        print("═══ ЖИВАЯ ВИНЬЕТКА «ЧТО НОВОГО» ═══")
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
