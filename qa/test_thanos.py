#!/usr/bin/env python3
"""Тест эффекта удаления («танос») после переработки на запечённые ролики.

Сценарий юзера: при удалении дежурств и компенсаций карточка растворяется
в пыли. Раньше холст эффекта был на всё окно и обновлялся каждый кадр —
это грузило слабые машины (перезаливка текстуры размером с окно), а у
карточек с дробной шириной (ширина ячейки сетки) выживала только верхняя
строка пыли. Теперь: холст по размеру карточки, траектории «запечены»
(10 вариантов на размер, генерируются один раз, дальше переигрываются).

Инварианты:
  - взрыв завершается: snapshotTaken и finished приходят, эффект
    освобождается, карточке возвращается opacity;
  - холст — карточка плюс запас на разлёт, а НЕ всё окно
    (площадь < 5% окна — защита от возврата к 960 000 px);
  - карточка с дробной шириной рождает столько же пыли, сколько целая;
  - запечённые варианты: пул на размер растёт до 10 и не дальше,
    повторные вызовы возвращают уже запечённые варианты (те же объекты).

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 python3 qa/test_thanos.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

WRAPPER = """
import QtQuick
import QtQuick.Window
import "components" as AppUI

Window {
    id: w
    width: 1200; height: 800; visible: true
    color: "#E4E7EA"

    // Чип дежурства: целая ширина
    Rectangle { id: card; objectName: "card"
        x: 300; y: 200; width: 130; height: 18; radius: 4; color: "#dbeafe"
        Text { anchors.centerIn: parent; text: "Дн 08:00–20:00"; font.pixelSize: 9; color: "#1e40af" } }
    // Чип с дробной шириной, как у реальных ячеек сетки (130.857)
    Rectangle { id: cardFrac; objectName: "cardFrac"
        x: 300; y: 260; width: 130.857; height: 18; radius: 4; color: "#dcfce7"
        Text { anchors.centerIn: parent; text: "Н 20:00–08:00"; font.pixelSize: 9; color: "#166534" } }
    // Бейдж компенсации 18x18
    Rectangle { id: badge; objectName: "badge"
        x: 360; y: 320; width: 18; height: 18; radius: 9; color: "#ccfbf1"
        Text { anchors.centerIn: parent; text: "В"; font.pixelSize: 9; color: "#0f766e" } }

    AppUI.ThanosEffect {
        id: effect
        objectName: "effect"
        anchors.fill: parent
    }

    // Повторы обязаны проигрывать уже запечённые варианты (те же объекты).
    // Проверяется внутри QML: через границу Python объекты теряют идентичность.
    function replayIdentity() {
        var key = "130x18"
        if (!effect.variantCache[key]) effect.takeVariant(130, 18)
        for (var i = effect.variantCache[key].length; i < 10; i++) effect.takeVariant(130, 18)
        var pool = effect.variantCache[key]
        if (pool.length !== 10) return [false, pool.length]
        for (var k = 0; k < 50; k++) {
            if (pool.indexOf(effect.takeVariant(130, 18)) < 0) return [false, -1]
        }
        return [effect.variantCache[key].length === 10, effect.variantCache[key].length]
    }
}
"""

WINDOW_AREA = 1200 * 800


def as_python(value):
    """var-свойства QML приходят в Python как QJSValue — разворачиваем."""
    try:
        return value.toVariant()
    except AttributeError:
        return value


def find_canvas(effect):
    for child in effect.childItems():
        if "Canvas" in child.metaObject().className():
            return child
    return None


def wait_finished(app, effect, seconds=15.0):
    deadline = time.time() + seconds
    while time.time() < deadline:
        app.processEvents()
        if not effect.property("isExploding"):
            return True
        time.sleep(0.01)
    return False


def main() -> int:
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuick import QQuickItem

    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    wrapper = os.path.join(ROOT, "_render_thanos.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "сцена не загрузилась"
        deadline = time.time() + 2.5
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.02)
        win = engine.rootObjects()[0]
        effect = win.findChild(QQuickItem, "effect")
        card = win.findChild(QQuickItem, "card")
        card_frac = win.findChild(QQuickItem, "cardFrac")
        badge = win.findChild(QQuickItem, "badge")
        assert effect and card and card_frac and badge, "эффект или карточки не найдены"

        # ── 1. Запечённые варианты: пул до 10, повторы — те же объекты ──
        ok, info = win.replayIdentity().toVariant()
        assert ok, f"повторы не играют запечённые варианты (info={info})"
        print(f"варианты: пул 130x18 = {info}, 50 повторов — те же запечённые объекты")

        # Сигналы и счёт пыли в момент снимка (по завершении массив уже очищен)
        seen = {"snap": 0, "fin": 0, "parts": 0}

        def on_snap():
            seen["snap"] += 1
            seen["parts"] = len(as_python(effect.property("particles")) or [])

        def on_fin():
            seen["fin"] += 1

        effect.snapshotTaken.connect(on_snap)
        effect.finished.connect(on_fin)

        # ── 2. Взрыв целого чипа: завершается, холст маленький ──
        effect.explode(card)
        canvas = find_canvas(effect)
        assert canvas is not None, "холст эффекта не найден"
        cw, chh = canvas.width(), canvas.height()
        cx, cy = canvas.x(), canvas.y()
        assert cw >= 130 and chh >= 18, f"холст меньше карточки: {cw}x{chh}"
        # 24/80/150/24 — запасы на разлёт из компонента (плюс допуск на округление)
        assert cw <= 130 + 24 + 150 + 2 and chh <= 18 + 80 + 24 + 2, \
            f"холст больше карточки с запасом: {cw}x{chh}"
        assert cx <= 300 and cx + cw >= 430, "карточка не внутри холста по X"
        assert cy <= 200 and cy + chh >= 218, "карточка не внутри холста по Y"
        area = cw * chh
        assert area < 0.05 * WINDOW_AREA, f"холст снова почти на всё окно: {area} px"
        assert wait_finished(app, effect), "взрыв целого чипа не завершился"
        assert seen["snap"] == 1 and seen["fin"] == 1, \
            f"сигналы: snap={seen['snap']} fin={seen['fin']}"
        assert card.opacity() == 1.0, "карточке не вернули opacity"
        n_int = seen["parts"]
        assert 200 <= n_int <= 380, f"пыли у целого чипа неожиданно мало/много: {n_int}"
        print(f"целый чип: холст {cw}x{chh} = {area} px (было 960 000), пылинок {n_int}")

        # ── 3. Дробная ширина: столько же пыли, сколько у целой ──
        seen.update(snap=0, fin=0, parts=0)
        effect.explode(card_frac)
        assert wait_finished(app, effect), "взрыв дробного чипа не завершился"
        n_frac = seen["parts"]
        assert n_frac >= 200, f"у дробной ширины пыль снова пропадает: {n_frac} (было 33)"
        ratio = n_frac / float(n_int)
        assert 0.7 <= ratio <= 1.4, f"пыль целой и дробной карточки расходится: {n_int} vs {n_frac}"
        assert card_frac.opacity() == 1.0, "дробному чипу не вернули opacity"
        print(f"дробный чип: пылинок {n_frac} (у целой {n_int}, отношение {ratio:.2f}; раньше было 33)")

        # ── 4. Бейдж 18x18 ──
        seen.update(snap=0, fin=0, parts=0)
        effect.explode(badge)
        canvas = find_canvas(effect)
        area_b = canvas.width() * canvas.height()
        assert wait_finished(app, effect), "взрыв бейджа не завершился"
        assert area_b < 0.05 * WINDOW_AREA, f"холст бейджа велик: {area_b} px"
        assert badge.opacity() == 1.0, "бейджу не вернули opacity"
        n_badge = seen["parts"]
        assert 15 <= n_badge <= 60, f"пыли у бейджа неожиданно мало/много: {n_badge}"
        print(f"бейдж: холст {canvas.width()}x{canvas.height()} = {area_b} px, пылинок {n_badge}")

        # ── 5. Пул вариантов не растёт дальше 10 ──
        pool = as_python(effect.property("variantCache")) or {}
        assert "130x18" in pool, "нет пула для 130x18"
        assert len(pool["130x18"]) == 10, f"пул 130x18 = {len(pool['130x18'])}, а не 10"
        print(f"пул вариантов 130x18: {len(pool['130x18'])} (лимит), размеров в кэше: {len(pool)}")

        print("═══ ТАНОС ЦЕЛ ═══")
        return 0
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
