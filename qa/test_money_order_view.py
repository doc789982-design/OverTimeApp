#!/usr/bin/env python3
"""Рендер и проверка окна «Ведомость» (MoneyOrderDialog) в песочнице.

После переделки окна по браку сборки 257 и доработок сборки 259
проверяем программно (тема песочницы бывает светлой ИЛИ тёмной —
все проверки обязаны работать в обеих):
  - список сотрудников НЕ пустой: три строки-КАРТОЧКИ (заливка
    bgCell, как числа месяца), у каждой слева имя, справа три
    компактных числовых поля;
  - позиции полей совпадают по строкам (подписи не разъехались);
  - в шапке таблицы ГАЛОЧКА «ВСЕ» (при всех выбранных — акцентный
    квадрат) и колонки заголовка выровнены с полями строк;
  - внизу итог «Выплата: …» и кнопки «Закрыть» / «Провести приказ».
Снимок: qa/render/money_order.png

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_money_order_view.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "qa", "render")

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 1200; height: 800; visible: true
    AppUI.MoneyOrderDialog { id: dlg }
    Component.onCompleted: dlg.openNew()
}
"""

EMPS = [
    {"id": 1, "name": "Иванов Иван", "subtitle": "капитан — инженер",
     "hours": 8, "overtime": 4, "days": 2,
     "prevHours": 5, "prevOvertime": 1, "prevDays": 1},
    {"id": 2, "name": "Петров Пётр Петрович", "subtitle": "лейтенант — инженер",
     "hours": 4, "overtime": 0, "days": 0,
     "prevHours": 2, "prevOvertime": 0, "prevDays": 0},
    {"id": 3, "name": "Сидоров Сидор", "subtitle": "",
     "hours": 0, "overtime": 0, "days": 1,
     "prevHours": 0, "prevOvertime": 0, "prevDays": 0},
]


from PySide6.QtCore import QObject, Slot
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine, QJSValue
from PySide6.QtQuick import QQuickItem  # регает конвертер QQuickItem*


class StubBackend(QObject):
    """Стаб на уровне модуля (локальные классы PySide6 регистрирует
    ненадёжно — окно рендерится как голый QWindow)."""

    @Slot(result="QVariant")
    def moneyOrderEmployees(self):
        return EMPS

    @Slot(str, str, str, str, int, result="QVariant")
    def saveMoneyOrder(self, payload, order_no, order_date_iso, comment,
                       source_mode):
        return {"ok": True, "count": 1}


def to_var(v):
    return v.toVariant() if isinstance(v, QJSValue) else v


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()          # ссылка обязательна (GC)
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_money.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)

    errors = []
    engine.warnings.connect(
        lambda ws: errors.extend(str(x.toString()) for x in ws))
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]

        # диалог — Popup (QObject, не айтем): ищем среди детей окна
        dlg = None
        for o in win.findChildren(QObject):
            if o.metaObject().className().startswith("MoneyOrderDialog"):
                dlg = o
                break
        assert dlg is not None, "диалог ведомости не найден"

        # даём раскладке построиться
        deadline = time.time() + 3
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.02)

        # 1. список сотрудников заполнен
        rows = to_var(dlg.property("rows"))
        assert len(rows) == 3, "rows не заполнен: %r" % (rows,)

        # 2. геометрия отрисованных строк: высоты, ширины, колонки
        info = to_var(dlg.layoutInfo())
        hs = [h for h in info["heights"] if h > 0]
        ws = [w for w in info["widths"] if w > 0]
        assert len(hs) == 3 and all(h >= 60 for h in hs), \
            "строки не отрисовались: %r" % (info,)
        assert all(abs(w - ws[0]) < 2 for w in ws), \
            "строки разной ширины: %r" % (ws,)
        assert abs(info["headerFieldX"] - info["rowFieldX"]) < 6, \
            "колонки заголовка и строк не выровнены: %r" % (info,)

        # 3. ширина колонки имени достаточна (поля не отжали имя)
        name_w = float(dlg.property("nameWidth"))
        assert name_w > 150, "колонка имени слишком узкая: %s" % name_w

        # 3б. структура окна: радио годов, журнал, поля слева,
        #     «Доступно» и «Выплата» убраны
        radios = [o for o in dlg.findChildren(QObject)
                  if o.metaObject().className().startswith("AppRadioButton")
                  and o.property("visible")]
        assert len(radios) == 3, "ожидали 3 радио-кнопки года: %d" % len(radios)
        rtexts = sorted(str(r.property("text")) for r in radios)
        assert rtexts == ["Оба года", "Предыдущий год", "Текущий год"], rtexts
        journal_btns = [o for o in dlg.findChildren(QObject)
                        if o.property("text") == "Журнал приказов"
                        and o.property("visible")]
        assert journal_btns, "нет кнопки «Журнал приказов»"
        # реквизиты приказа: дата → номер → комментарий, вертикально слева
        fields = [o for o in dlg.findChildren(QObject)
                  if o.metaObject().className().startswith(("AppTextField",
                                                            "AppDateField"))]
        head = [f for f in fields
                if str(f.property("label") or "") in
                ("Дата приказа", "Номер приказа", "Комментарий")]
        assert len(head) == 3, "нет трёх полей реквизитов: %d" % len(head)
        for f in head:
            assert float(f.property("x") or 0) < 40, \
                "поле не у левого края: %s x=%s" % (
                    f.property("label"), f.property("x"))
        ys = sorted((float(f.property("y") or 0),
                     str(f.property("label"))) for f in head)
        assert [t for _, t in ys] == ["Дата приказа", "Номер приказа",
                                      "Комментарий"], ys
        # «Доступно» и «Выплата» больше нет
        texts = []
        for o in dlg.findChildren(QObject):
            t = o.property("text")
            if t and o.property("visible"):
                texts.append(str(t))
        assert not any("Доступно" in t for t in texts), "«Доступно» не убрано"
        assert not any("Выплата" in t for t in texts), "«Выплата» не убрано"

        # 3в. радио годов не обрезаны многоточием (сборка 262:
        #     implicitWidth забыл x индикатора → вечное «…»)
        for r in radios:
            lbl = None
            for c in r.findChildren(QObject):
                if c.metaObject().className() == "QQuickText":
                    lbl = c
                    break
            assert lbl is not None, "у радио нет лейбла"
            assert lbl.property("truncated") is not True, \
                "радио %r обрезано: %r" % (r.property("text"),
                                           lbl.property("text"))

        # 3г. поля сумм НЕ выезжают за правый край карточки
        #     (nameWidth теперь от реальной ширины таблицы).
        #     Делегаты Repeater ищем через childItems — QObject-
        #     findChildren их не видит
        def walk_items(item, acc):
            acc.append(item)
            for c in item.childItems():
                walk_items(c, acc)
            return acc

        all_items = []
        for root_i in win.findChildren(QQuickItem):
            if root_i.parentItem() is None:
                walk_items(root_i, all_items)
        cards = [i for i in all_items
                 if i.metaObject().className().startswith("QQuickRectangle")
                 and abs(float(i.property("height") or 0) - 66) < 1]
        assert cards, "карточки-строки не найдены"
        card = cards[0]
        card_w = float(card.property("width"))
        card_items = walk_items(card, [])
        fields_in_card = [i for i in card_items
                          if i.metaObject().className().startswith("AppTextField")
                          and abs(float(i.property("width") or 0) - 56) < 1]
        assert len(fields_in_card) == 3, \
            "в карточке нет трёх полей сумм: %d" % len(fields_in_card)
        for f in fields_in_card:
            right = float(f.property("x") or 0) + 56
            assert right <= card_w - 8 + 0.5, \
                "поле выезжает за карточку: правый край %.1f при ширине %.1f" \
                % (right, card_w)

        # 4. снимок карточки (contentItem попапа — «окно» ведомости).
        #    В offscreen раскладка/заливки «доезжают» не сразу: первые
        #    grabToImage — прогрев, боевой снимок — последний.
        ci = dlg.property("contentItem")
        assert ci is not None and isinstance(ci, QQuickItem), "contentItem?"

        def grab():
            holder = []
            res = ci.grabToImage()
            assert res is not None, "grabToImage не удался"
            res.ready.connect(lambda: holder.append(res.image()))
            t0 = time.time()
            while not holder and time.time() - t0 < 5:
                app.processEvents()
                time.sleep(0.02)
            assert holder, "снимок не пришёл"
            return holder[0]

        for i in range(6):
            grab()
            time.sleep(0.4)
            for _ in range(10):
                app.processEvents()
        img = grab()                # боевой снимок
        shot = os.path.join(OUT, "money_order.png")
        img.save(shot)
        assert img.width() >= 600 and img.height() >= 300, \
            "карточка подозрительно мала: %dx%d" % (img.width(), img.height())

        # 5. пиксельная структура снимка. Поля AppTextField прозрачны,
        #    рисуется рамка 1px (borderInput ≈ (199,205,209)); строки —
        #    карточки с заливкой bgCell (тёмная (36,36,38) / светлая
        #    (244,245,247) — тема песочницы непостоянна).
        from PIL import Image
        pic = Image.open(shot).convert("RGB")
        W, H = pic.size
        px = pic.load()

        def near(p, ref, tol=30):
            return abs(p[0]-ref[0]) + abs(p[1]-ref[1]) + abs(p[2]-ref[2]) < tol

        BORDER = (199, 205, 209)
        ACCENT = (3, 116, 181)
        CELL_DARK = (36, 36, 38)
        CELL_LIGHT = (244, 245, 247)

        # тему сэмплируем по центру первой строки (между рамками полей)
        cell = CELL_DARK
        for y in range(H // 3, H * 2 // 3):
            nd = sum(1 for x in range(40, W - 40, 4)
                     if near(px[x, y], CELL_DARK))
            nl = sum(1 for x in range(40, W - 40, 4)
                     if near(px[x, y], CELL_LIGHT))
            if nd > 60 or nl > 60:
                cell = CELL_DARK if nd > nl else CELL_LIGHT
                break

        # 5а. строки-карточки: горизонтальные полосы заливки bgCell
        bands = []
        in_b = False
        for y in range(H):
            n = sum(1 for x in range(0, W, 4) if near(px[x, y], cell))
            if n > 100 and not in_b:
                in_b, y0 = True, y
            elif n <= 100 and in_b:
                in_b = False
                if y - y0 > 45:
                    bands.append((y0, y))
        if in_b and H - y0 > 45:
            bands.append((y0, H))
        assert len(bands) == 3, \
            "ожидали 3 карточки-строки (bgCell), нашли %d: %r" % (len(bands), bands)
        bh = [b[1] - b[0] for b in bands]
        assert max(bh) - min(bh) <= 2, "карточки разной высоты: %r" % bh
        for (by0, by1) in bands:
            ym = (by0 + by1) // 2
            xs = [x for x in range(W) if near(px[x, ym], cell)]
            assert max(xs) - min(xs) > 500, \
                "карточка y=%d уже 500px (%d..%d)" % (ym, min(xs), max(xs))

        # 5б. рамки полей: y -> x-прогоны цвета рамки ≥ 36px
        lines = {}
        for y in range(H):
            runs, in_r = [], False
            for x in range(W):
                ok = near(px[x, y], BORDER)
                if ok and not in_r:
                    in_r, rx = True, x
                elif not ok and in_r:
                    in_r = False
                    if x - rx >= 36:
                        runs.append((rx, x))
            if in_r and W - rx >= 36:
                runs.append((rx, W))
            if runs:
                lines[y] = runs

        # строки таблицы: y с ровно тремя прогонами (три числовых поля)
        table_rows = {y: r for y, r in lines.items()
                      if len(r) == 3 and y > H * 0.4}
        assert len(table_rows) >= 2, \
            "не нашли строк таблицы с полями: %r" % (list(lines.items())[:8],)
        ys = sorted(table_rows)
        stripes = []
        i = 0
        while i < len(ys):
            group = [ys[i]]
            while i + 1 < len(ys) and ys[i+1] - ys[i] <= 76:
                same = all(
                    abs(a[0]-b[0]) <= 3 and abs(a[1]-b[1]) <= 3
                    for a, b in zip(table_rows[group[0]], table_rows[ys[i+1]]))
                if not same:
                    break
                group.append(ys[i+1])
                i += 1
            if len(group) >= 3:
                stripes.append(group)
            i += 1
        assert len(stripes) >= 1, \
            "нет трёх строк с совпадающими колонками: %r" % (ys,)
        cols = table_rows[stripes[0][0]]

        # поля компактные: ширина 40..75 — НЕ растянуты по горизонтали
        fws = [c[1] - c[0] for c in cols]
        assert all(40 <= w <= 75 for w in fws), \
            "поля растянуты по горизонтали: %r" % fws
        for y in stripes[0][1:]:
            for a, b in zip(cols, table_rows[y]):
                assert abs(a[0]-b[0]) <= 3 and abs(a[1]-b[1]) <= 3, \
                    "колонки разъехались: %r vs %r" % (a, b)

        # в зоне таблицы нет ГОРИЗОНТАЛЬНО растянутых полей (>150px)
        y0t, y1t = min(stripes[0]) - 10, max(stripes[0]) + 10
        wide = [(y, r) for y, r in lines.items()
                if y0t <= y <= y1t and any(b - a > 150 for a, b in r)]
        assert not wide, "растянутые поля в таблице: %r" % (wide,)

        # 5в. имена слева: акцентный текст в каждой из трёх строк
        for y in stripes[0]:
            n = sum(1 for x in range(10, cols[0][0] - 10)
                    for yy in range(y - 55, y - 15)
                    if near(px[x, yy], ACCENT, 60))
            assert n > 30, "имя сотрудника не найдено в строке y=%d" % y

        # 5г. галочка «все» в заголовке таблицы: акцентный квадрат
        #     (после openNew все выбраны → галочка включена)
        n_master = sum(1 for x in range(15, 60) for y in range(335, 380)
                       if near(px[x, y], ACCENT, 40))
        assert n_master > 150, \
            "галочка «все» в заголовке не найдена (%d px)" % n_master

        # 5д. радио «Текущий год» отмечено: акцентная точка-индикатор
        #     в правой колонке шапки
        n_radio = sum(1 for x in range(255, 305) for y in range(115, 165)
                      if near(px[x, y], ACCENT, 40))
        assert n_radio > 40, "радио года не найдено (%d px)" % n_radio

        # 6. кнопки читаемы (тема-агностично; итог «Выплата» убран)
        def opaque(y0, y1, x0, x1):
            return sum(1 for x in range(x0, x1) for y in range(y0, y1)
                       if px[x, y] != (0, 0, 0))

        n_cancel = opaque(628, 685, 195, 310)
        assert n_cancel > 80, "текст «Закрыть» не читаем (%d)" % n_cancel
        save_fill = sum(1 for x in range(290, 460)
                        for y in range(628, 685)
                        if near(px[x, y], ACCENT, 40))
        assert save_fill > 2000, "кнопка «Провести приказ» не найдена"

        # 7. ошибок QML в диалоге нет
        bad = [e for e in errors
               if "not a function" in e or "Cannot assign" in e
               or "read-only" in e or "ReferenceError" in e
               or ("TypeError" in e and "isDarkTheme" not in e)]
        assert not bad, "ошибки QML: %r" % (bad[:3],)

        print("Ведомость: поля слева вертикально, радио года и галочка «все»"
              " на месте, 3 карточки-строки %r, поля %r px, кнопки читаемы"
              " — снимок %s" % (bh, fws, shot))
        return 0
    finally:
        if os.path.exists(wrapper):
            os.remove(wrapper)


if __name__ == "__main__":
    sys.exit(main())
