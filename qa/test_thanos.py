#!/usr/bin/env python3
"""Тест эффекта удаления («танос»): два пути — GPU-шейдер и запасной Canvas.

Сценарий юзера: при удалении дежурств и компенсаций карточка растворяется
в пыли. Идея шейдерного пути взята из Telegram Desktop (thanos_effect):
случайность не хранят — вычисляют хешем от координаты и зерна, цвет
пылинка берёт из самого снимка карточки; CPU в кадре не считает ничего.
Пылинка — мягкая капля, вытянутая по полёту и сжимающаяся к концу жизни
(не пиксельный шум). На машинах без графического ускорителя (и в
песочнице тестов — софтверный рендер) работает запасной путь: Canvas
по размеру карточки с запечёнными вариантами траекторий.

Инварианты:
  - под софтверным рендером эффект идёт запасным путём: шейдерные слои
    даже не создаются (Loader неактивен), GraphicsInfo видит Software;
  - взрыв завершается: snapshotTaken и finished приходят, эффект
    освобождается, карточке возвращается opacity;
  - холст — карточка плюс запас на разлёт, а НЕ всё окно
    (площадь < 5% окна — защита от возврата к 960 000 px);
  - карточка с дробной шириной рождает столько же пыли, сколько целая;
  - запечённые варианты: пул на размер растёт до 10 и не дальше,
    повторные вызовы возвращают уже запечённые варианты (те же объекты);
  - шейдер на месте и «дорогой»: контракт Qt (qt_TexCoord0, qt_Matrix
    первым в блоке, сэмплер на binding 1, premultiplied), мягкий край
    (smoothstep), сжатие капли к концу жизни, кольцо соседних ячеек 5x5
    покрывает радиус капли с вытяжкой, у каждого слоя своё зерно;
  - математика пыли (зеркало формул на Python): пыль гаснет к концу
    эффекта, не вылетает за запас холста, и её ДОСТАТОЧНО — капель
    заметно больше сотни, покрытие кратно больше пиксельного вида.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 python3 qa/test_thanos.py
"""
import math
import os
import random
import re
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


def parse_list(qml_text, name):
    m = re.search(rf"{name}:\s*\[([^\]]+)\]", qml_text)
    assert m, f"в ThanosEffect.qml нет списка {name}"
    return [float(x) for x in m.group(1).split(",")]


def check_shader_files():
    """Контракт шейдера, «дорогой» вид и упаковка в exe."""
    frag_path = os.path.join(ROOT, "shaders", "thanos_dust.frag")
    qsb_path = frag_path + ".qsb"
    assert os.path.exists(frag_path), "нет shaders/thanos_dust.frag"
    assert os.path.exists(qsb_path), "нет запечённого shaders/thanos_dust.frag.qsb"
    frag = open(frag_path, encoding="utf-8").read()
    effect_qml = open(os.path.join(ROOT, "components", "ThanosEffect.qml"),
                      encoding="utf-8").read()

    # Контракт Qt для fragment-only шейдера ShaderEffect (дока Qt)
    assert "in vec2 qt_TexCoord0;" in frag, "вход обязан называться qt_TexCoord0"
    m = re.search(r"uniform buf \{\s*mat4 qt_Matrix;\s*float qt_Opacity;", frag)
    assert m, "блок обязан начинаться с qt_Matrix и qt_Opacity (дока Qt)"
    assert "layout(binding = 1) uniform sampler2D src;" in frag, "сэмплер на binding 1"
    assert "fragColor = vec4(accRGB, accA);" in frag, "выход premultiplied"
    print("шейдер: контракт Qt соблюдён (qt_TexCoord0, qt_Matrix первым, binding 1, premultiplied)")

    # «Дорогой» вид: мягкая капля, сжатие к концу жизни, плотность
    assert "smoothstep(0.35, 1.0, length(q))" in frag, "у капли нет мягкого края"
    assert "(0.35 + 0.65 * a)" in frag, "капля не сжимается к концу жизни"
    assert "for (int j = -2; j <= 2; j++)" in frag, "кольцо соседних ячеек обязано быть 5x5"
    assert "for (int s = 0; s < 2; s++)" in frag, "в ячейке нет второй пылинки"
    assert "clamp(uClass.x - float(s), 0.0, 1.0)" in frag, "нет ворот плотности по подп слотам"
    assert "root.seed + index" in effect_qml, "зерно слоя не сдвинуто по индексу"
    sizes = parse_list(effect_qml, "dustSizes")
    densities = parse_list(effect_qml, "dustDensities")
    assert len(sizes) == 3 and len(densities) == 3, "нужны три семьи пыли"
    assert all(2.0 <= s <= 6.0 for s in sizes), f"размеры капель вне рынка: {sizes}"
    assert sum(densities) >= 3.5, f"пыли мало: суммарная плотность {sum(densities)}"
    print(f"вид: мягкие капли {sizes} px, плотность на ячейку {densities}, "
          "сжатие к концу жизни, зерно слоя своё")

    # .qsb — сжатый QShader: 4 байта размера (big-endian) + zlib-поток
    head = open(qsb_path, "rb").read(6)
    size = int.from_bytes(head[:4], "big")
    assert size > 1000, f"подозрительно маленький .qsb: {size} байт"
    assert head[4:6] == b"\x78\x9c", "после размера ждём zlib-поток (сжатый QShader)"
    print(f".qsb: валиден (сжатый QShader, {size} байт)")

    # Упаковка в exe: папка shaders и расширение .qsb в генераторе ресурсов
    mr = open(os.path.join(ROOT, "tools", "make_resources.py"), encoding="utf-8").read()
    assert re.search(r'DIRS\s*=\s*\[[^\]]*"shaders"', mr), "shaders не в DIRS генератора ресурсов"
    assert '".qsb"' in mr, ".qsb не в INCLUDE_EXT генератора ресурсов"
    print("упаковка: shaders/ и .qsb едут в exe (make_resources)")
    return frag, effect_qml


def check_math(frag, effect_qml):
    """Зеркало формул шейдера: границы, насыщенность, кольцо покрытия."""
    motions = [
        tuple(float(g) for g in m)
        for m in re.findall(r"Qt\.vector4d\(([-\d.]+),\s*([-\d.]+),\s*([-\d.]+),\s*([-\d.]+)\)",
                            effect_qml)
    ]
    assert len(motions) >= 2, "в эффекте должно быть несколько семей пыли"
    sizes = parse_list(effect_qml, "dustSizes")
    densities = parse_list(effect_qml, "dustDensities")
    size_jitter = float(re.search(r"dustSizeJitter:\s*([\d.]+)", effect_qml).group(1))
    life_m = re.search(r"uLife:\s*Qt\.vector2d\(([\d.]+),\s*([\d.]+)\)", effect_qml)
    life_base, life_jitter = float(life_m.group(1)), float(life_m.group(2))
    wave = float(re.search(r"waveSec:\s*([\d.]+)", effect_qml).group(1))
    gravity = float(re.search(r"dustGravity:\s*([\d.]+)", effect_qml).group(1))
    k = float(re.search(r"float kDecay = ([\d.]+);", frag).group(1))
    cell = float(re.search(r"float cell = ([\d.]+);", frag).group(1))
    stretch = float(re.search(r"q /= vec2\(r \* ([\d.]+), r \* ([\d.]+)\);", frag).group(1))
    pads = {
        p: float(re.search(rf"property real pad{p}:\s*([\d.]+)", effect_qml).group(1))
        for p in ("Left", "Top", "Right", "Bottom")
    }

    # кольцо 5x5 покрывает радиус капли: центр ячейки ±(сдвиг + вытяжка)
    # обязан укладываться в полутора ячейках — иначе капли режутся по решётке
    reach = (max(sizes) + size_jitter) * stretch + 3.0
    assert reach <= 1.5 * cell + 0.5, \
        f"капля достаёт на {reach:.1f} px, а кольцо 5x5 покрывает {1.5 * cell:.0f} px"

    def wake(th):
        return wave * (1.0 - math.sqrt(max(0.0, 1.0 - th)))

    total_m = re.search(r"totalSec:\s*waveSec \+ ([\d.]+) \+ ([\d.]+) \+ ([\d.]+)", effect_qml)
    total = wave + float(total_m.group(1)) + float(total_m.group(2)) + float(total_m.group(3))
    max_life = life_base + life_jitter
    assert wake(1.0) + max_life <= total + 1e-6, \
        f"пыль живёт дольше эффекта: {wake(1.0) + max_life:.3f} > {total:.3f}"

    random.seed(7)
    xs, ys = [], []
    for (vx, vy, dx, dy) in motions:
        for _ in range(15000):
            sp = 0.6 + random.random() * 0.8
            ang = (random.random() - 0.5) * 0.7
            ca, sa = math.cos(ang), math.sin(ang)
            rvx = (ca * vx - sa * vy) * sp
            rvy = (sa * vx + ca * vy) * sp
            life = life_base + random.random() * life_jitter
            tau = min(max_life, life * 0.95)
            e = 1.0 - math.exp(-k * tau)
            ox = rvx / k * e + dx * tau
            oy = rvy / k * e + dy * tau + 0.5 * gravity * tau * tau
            xs.append(ox)
            ys.append(oy)
    xs.sort()
    ys.sort()

    def pct(a, p):
        return a[min(len(a) - 1, int(len(a) * p))]

    assert pct(xs, 0.99) <= pads["Right"] + 2, \
        f"пыль вылетает вправо: 99-й перцентиль {pct(xs, 0.99):.0f} > запаса {pads['Right']}"
    assert pct(ys, 0.01) >= -(pads["Top"] + 2), \
        f"пыль вылетает вверх: {pct(ys, 0.01):.0f} < запаса {pads['Top']}"
    assert xs[0] >= -(pads["Left"] + 2), f"пыль вылетает влево: {xs[0]:.0f}"
    assert ys[-1] <= pads["Bottom"] + 2, f"пыль вылетает вниз: {ys[-1]:.0f}"
    print(f"математика: гаснет к {total:.2f} с; разлёт X 99% ≤ {pct(xs, 0.99):.0f} px "
          f"(запас {pads['Right']:.0f}), вверх ≤ {abs(pct(ys, 0.01)):.0f} px (запас {pads['Top']:.0f})")

    # Насыщенность: мини-симуляция сетки в разгар эффекта — сколько капель живо
    card_w, card_h = 130, 18
    for t in (0.5, 1.0):
        alive = 0
        for cls in range(3):
            seed = 0.42 + cls * 0.7371

            def h(x, y):
                return (math.sin(x * 127.1 + y * 311.7 + seed * 17.13) * 43758.5453) % 1.0

            vx, vy, dx, dy = motions[cls]
            ncy = int((card_h + pads["Top"] + pads["Bottom"]) / cell) + 3
            ncx = int((card_w + pads["Left"] + pads["Right"]) / cell) + 3
            for cy in range(-2, ncy):
                for cx in range(-2, ncx):
                    for s in range(2):
                        gate = min(1.0, max(0.0, densities[cls] - s))
                        if h(cx + s * 0.37, cy + s * 0.71) > gate:
                            continue
                        px_ = (cx + 0.5) * cell + (h(cx + s * 0.37 + 11.7, cy + s * 0.71 + 21.3) - 0.5) * 6.0
                        py_ = (cy + 0.5) * cell + (h(cx + s * 0.37 + 19.1, cy + s * 0.71 + 8.7) - 0.5) * 6.0
                        fx = min(1.0, max(0.0, (px_ - pads["Left"]) / card_w))
                        th = (1.0 - fx) * 0.7 + h(cx + s * 0.37 + 23.0, cy + s * 0.71 + 5.0) * 0.3
                        tau = t - wake(th)
                        if tau <= 0:
                            continue
                        sp = 0.6 + h(cx + s * 0.37 + 7.7, cy + s * 0.71 + 3.1) * 0.8
                        ang = (h(cx + s * 0.37 + 5.3, cy + s * 0.71 + 9.2) - 0.5) * 0.7
                        ca, sa = math.cos(ang), math.sin(ang)
                        rvx = (ca * vx - sa * vy) * sp
                        rvy = (sa * vx + ca * vy) * sp
                        e = 1.0 - math.exp(-k * tau)
                        ox = rvx / k * e + dx * tau + math.sin(tau * 9.0 + ang * 6.2832) * 2.5
                        oy = (rvy / k * e + dy * tau + 0.5 * gravity * tau * tau
                              + math.cos(tau * 11.0 + sp * 6.2832) * 2.5)
                        ux = (px_ - ox - pads["Left"]) / card_w
                        uy = (py_ - oy - pads["Top"]) / card_h
                        if not (0.0 <= ux < 1.0 and 0.0 <= uy < 1.0):
                            continue
                        life = life_base + h(cx + s * 0.37 + 13.9, cy + s * 0.71 + 27.4) * life_jitter
                        if tau < life:
                            alive += 1
        if t == 0.5:
            peak = alive
        else:
            mid = alive
    assert 120 <= peak <= 420, f"капель в разгар неожиданно мало/много: {peak}"
    assert mid >= 0.35 * peak, f"пыль гаснет слишком рано: {mid} из {peak}"
    print(f"насыщенность: в разгар {peak} капель, к середине {mid} — "
          f"в {(peak * math.pi * 1.44 * (sum(sizes)/3)**2) / 2340:.1f}× площади карточки")


def main() -> int:
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuick import QQuickItem

    frag, effect_qml = check_shader_files()
    check_math(frag, effect_qml)

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

        # ── 2. Под софтверным рендером — запасной путь ──
        assert effect.property("gpuOK") is False, \
            "в песочнице (софтверный рендер) gpuOK обязан быть ложью"

        def walk(item, out):
            for c in item.childItems():
                out.append(c)
                walk(c, out)

        all_items = []
        walk(effect, all_items)
        loaders = [c for c in all_items if "Loader" in c.metaObject().className()]
        assert loaders, "Loader пыли не найден"
        assert loaders[0].property("active") is False, \
            "под софтверным рендером шейдерные слои не должны создаваться"
        print("ветвление: софтверный рендер → gpuOK ложь, шейдерные слои не создаются")

        # Сигналы и счёт пыли в момент снимка (по завершении массив уже очищен)
        seen = {"snap": 0, "fin": 0, "parts": 0}

        def on_snap():
            seen["snap"] += 1
            seen["parts"] = len(as_python(effect.property("particles")) or [])

        def on_fin():
            seen["fin"] += 1

        effect.snapshotTaken.connect(on_snap)
        effect.finished.connect(on_fin)

        # ── 3. Взрыв целого чипа: завершается, холст маленький ──
        effect.explode(card)
        canvas = find_canvas(effect)
        assert canvas is not None, "холст запасного пути не найден"
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

        # ── 4. Дробная ширина: столько же пыли, сколько у целой ──
        seen.update(snap=0, fin=0, parts=0)
        effect.explode(card_frac)
        assert wait_finished(app, effect), "взрыв дробного чипа не завершился"
        n_frac = seen["parts"]
        assert n_frac >= 200, f"у дробной ширины пыль снова пропадает: {n_frac} (было 33)"
        ratio = n_frac / float(n_int)
        assert 0.7 <= ratio <= 1.4, f"пыль целой и дробной карточки расходится: {n_int} vs {n_frac}"
        assert card_frac.opacity() == 1.0, "дробному чипу не вернули opacity"
        print(f"дробный чип: пылинок {n_frac} (у целой {n_int}, отношение {ratio:.2f}; раньше было 33)")

        # ── 5. Бейдж 18x18 ──
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

        # ── 6. Пул вариантов не растёт дальше 10 ──
        pool = as_python(effect.property("variantCache")) or {}
        assert "130x18" in pool, "нет пула для 130x18"
        assert len(pool["130x18"]) == 10, f"пул 130x18 = {len(pool['130x18'])}, а не 10"
        print(f"пул вариантов 130x18: {len(pool['130x18'])} (лимит), размеров в кэше: {len(pool)}")

        print("═══ ТАНОС ЦЕЛ ═══")
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass

    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    # PySide6 6.12 падает при финализации интерпретатора после QML-сцены с
    # Canvas (none_dealloc на выходе, даже с нетронутым тестом прошлой
    # сборки) — выходим сразу после проверок, код возврата честный
    os._exit(rc)
