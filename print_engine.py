# -*- coding: utf-8 -*-
"""
Собственный движок печати бланка — без Excel.

Логика: заполненный шаблон (openpyxl, как раньше) отрисовывается напрямую
на QPrinter через QPainter. Вся геометрия — ширины колонок, высоты строк,
объединённые ячейки, границы, заливки, шрифты, выравнивания, поворот
текста — читается из самого xlsx, поэтому бланк печатается таким, каким
его видит Excel, без запуска Excel.

Возможности:
  - печать на выбранный принтер (копии, диапазон страниц, дублет);
  - разбивка на страницы по высоте, как в Excel;
  - выгрузка в PDF и картинка (превью/тесты) тем же кодом.
"""
from __future__ import annotations

import datetime
import os
import re
import tempfile
import uuid
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from PySide6.QtCore import QMarginsF, QRectF
from PySide6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter, QPen, QTransform
from PySide6.QtPrintSupport import QPrinter

MDW = 7            # ширина цифры шрифта Normal (Calibri 11) в px при 96 dpi
DEFAULT_COL_CHARS = 8.43
PX_PER_PT = 96.0 / 72.0

# подмена шрифтов: приоритетные кандидаты (первый найденный в системе)
FONT_CHAINS = {
    "PT Astra Serif": ["PT Astra Serif", "PT Astra Serif Ext", "Times New Roman", "Liberation Serif"],
    "Calibri": ["Calibri", "Carlito", "Segoe UI", "Liberation Sans"],
}

_DAY_FORMULA = re.compile(r'^=IFERROR\(DAY\(([A-Z]{1,3})(\d+)\),')


def _col_chars_to_px(chars: float) -> int:
    return int(round(chars * MDW)) + 5


class SheetModel:
    """Геометрия и стили листа в «листовых пикселях» (96 dpi)."""

    def __init__(self, ws):
        self.ws = ws
        # ── область печати ──
        pa = ws.print_area
        if pa:
            s = (pa if isinstance(pa, str) else pa[0]).split(",")[0].split("!")[-1].replace("$", "")
            cr = CellRange(s)
            self.c0, self.c1 = int(cr.min_col), int(cr.max_col)
            self.r0, self.r1 = int(cr.min_row), int(cr.max_row)
        else:
            self.c0, self.r0 = 1, 1
            self.c1, self.r1 = ws.max_column, ws.max_row

        # ── колонки: диапазоны <col min max> ──
        widths = {}
        for dim in ws.column_dimensions.values():
            if dim.width is None:
                continue
            lo = int(dim.min or 0)
            hi = int(dim.max or lo)
            for c in range(lo, hi + 1):
                widths[c] = (dim.width, bool(dim.hidden))
        default_w = ws.sheet_format.defaultColWidth or DEFAULT_COL_CHARS
        self.col_w = {}
        self.hidden_cols = set()
        for c in range(self.c0, self.c1 + 1):
            if c in widths:
                w, hidden = widths[c]
            else:
                w, hidden = default_w, False
            if hidden:
                self.col_w[c] = 0
                self.hidden_cols.add(c)
            else:
                self.col_w[c] = _col_chars_to_px(w)

        # ── ряды ──
        default_h = ws.sheet_format.defaultRowHeight or 15.0
        self.row_h = {}
        for r in range(self.r0, self.r1 + 1):
            d = ws.row_dimensions.get(r)
            h = (d.height if d and d.height else default_h) * PX_PER_PT
            self.row_h[r] = max(1, int(round(h)))

        # ── координаты ──
        self.x = {self.c0 - 1: 0}
        for c in range(self.c0, self.c1 + 1):
            self.x[c] = self.x[c - 1] + self.col_w[c]
        self.y = {self.r0 - 1: 0}
        for r in range(self.r0, self.r1 + 1):
            self.y[r] = self.y[r - 1] + self.row_h[r]

        self.width = self.x[self.c1]
        self.height = self.y[self.r1]

        # ── мержи (только внутри области печати) ──
        self.merges = {}
        for mr in ws.merged_cells.ranges:
            if mr.min_col > self.c1 or mr.max_col < self.c0 or mr.min_row > self.r1 or mr.max_row < self.r0:
                continue
            self.merges[(mr.min_row, mr.min_col)] = (mr.min_row, mr.max_row, mr.min_col, mr.max_col)
        self.merge_of = {}
        for a, (r1, r2, c1, c2) in self.merges.items():
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self.merge_of[(r, c)] = a

        # ── параметры страницы ──
        ps = ws.page_setup
        try:
            self.scale = (float(ps.scale) if ps.scale else 100.0) / 100.0
        except Exception:
            self.scale = 1.0
        m = ws.page_margins
        self.margins_mm = (25.4 * (m.left or 0.7), 25.4 * (m.right or 0.7),
                           25.4 * (m.top or 0.75), 25.4 * (m.bottom or 0.75))

    # ---------- содержимое ----------

    def cell_rect(self, r, c):
        a = self.merge_of.get((r, c))
        if a:
            r1, r2, c1, c2 = self.merges[a]
            return QRectF(self.x[c1 - 1], self.y[r1 - 1],
                          self.x[c2] - self.x[c1 - 1], self.y[r2] - self.y[r1 - 1])
        return QRectF(self.x[c - 1], self.y[r - 1], self.col_w[c], self.row_h[r])

    def is_anchor(self, r, c):
        a = self.merge_of.get((r, c))
        return a is None or a == (r, c)

    def value_for_print(self, cell):
        """Значение ячейки, как его показывает Excel на печати."""
        nf = (cell.number_format or "").strip()
        if nf == ";;;":          # формат «скрыть»
            return None
        v = cell.value
        if v is None:
            return None
        if isinstance(v, str):
            if v.startswith("="):
                # формулы шаблона скрыты форматом ;;; и не вычисляются;
                # на всякий случай поддержан единственный встречающийся вид
                m = _DAY_FORMULA.match(v)
                if m and nf != ";;;":
                    ref = self.ws[f"{m.group(1)}{m.group(2)}"].value
                    if isinstance(ref, (datetime.datetime, datetime.date)):
                        return str(ref.day)
                return None
            return v
        if isinstance(v, (datetime.datetime, datetime.date)):
            low = nf.lower()
            if ("d" in low and "m" not in low and "y" not in low) or low in ("general", "0"):
                return str(v.day)
            return v.strftime("%d.%m.%Y")
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)

    @staticmethod
    def color_of(cell, default=None):
        try:
            fill = cell.fill
            if fill and fill.fill_type == "solid":
                rgb = fill.start_color.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    return QColor("#" + rgb[-6:]), True
                th = getattr(fill.start_color, "theme", None)
                if th == 0:  # background1 — белый лист
                    return QColor("#ffffff"), True
        except Exception:
            pass
        return (default, False)


class SheetRenderer:
    """Рисует SheetModel на QPainter в листовых пикселях."""

    PEN_HAIR = 0.33    # 0.25 pt в листовых px
    PEN_THIN = 1.0     # 0.75 pt

    def __init__(self, model: SheetModel, font_hook=None):
        self.m = model
        self.font_hook = font_hook or (lambda name: name)   # для тестов: подмена шрифта
        self._fnt = {}

    def qfont(self, cell):
        f = cell.font
        name = f.name or "Calibri"
        chain = FONT_CHAINS.get(name)
        if chain:
            for cand in chain:
                fname = self.font_hook(cand)
                if fname:
                    name = fname
                    break
            else:
                name = self.font_hook(chain[-1]) or name
        else:
            name = self.font_hook(name) or name
        size = max(6, int(round((f.sz or 11.0) * PX_PER_PT)))
        key = (name, size, bool(f.b))
        if key not in self._fnt:
            font = QFont(name)
            font.setPixelSize(size)
            font.setBold(bool(f.b))
            font.setItalic(bool(f.i))
            self._fnt[key] = font
        return self._fnt[key]

    # ---------- страницы ----------

    def page_rows(self, page_h_px):
        """Разбивает видимые строки print_area на страницы по высоте."""
        pages, cur, cur_h = [], [], 0.0
        for r in range(self.m.r0, self.m.r1 + 1):
            h = self.m.row_h[r]
            if cur and cur_h + h > page_h_px:
                pages.append(cur)
                cur, cur_h = [], 0.0
            cur.append(r)
            cur_h += h
        if cur:
            pages.append(cur)
        return pages

    # ---------- отрисовка ----------

    def paint(self, painter: QPainter, rows, x_shift=0.0):
        m = self.m
        painter.save()
        painter.translate(-x_shift, -m.y[rows[0] - 1])

        # 1) заливки
        for r in rows:
            for c in range(m.c0, m.c1 + 1):
                if c in m.hidden_cols:
                    continue
                cell = m.ws.cell(r, c)
                color, ok = m.color_of(cell)
                if ok and color is not None:
                    rect = m.cell_rect(r, c)
                    painter.fillRect(rect, color)

        # 2) границы
        for r in rows:
            for c in range(m.c0, m.c1 + 1):
                if c in m.hidden_cols:
                    continue
                self._borders(painter, r, c)

        # 3) текст
        for r in rows:
            for c in range(m.c0, m.c1 + 1):
                if c in m.hidden_cols or not m.is_anchor(r, c):
                    continue
                cell = m.ws.cell(r, c)
                text = m.value_for_print(cell)
                if text is None or text == "":
                    continue
                self._text(painter, r, c, cell, str(text))
        painter.restore()

    def _borders(self, painter, r, c):
        m = self.m
        cell = m.ws.cell(r, c)
        a = m.merge_of.get((r, c))
        mr = m.merges.get(a) if a else None
        rect = m.cell_rect(r, c)

        def pen(side):
            b = getattr(cell.border, side, None)
            if not b or not b.style:
                return None
            color = "#808080" if b.style == "hair" else "#1a1a1a"
            try:
                if b.color and isinstance(b.color.rgb, str) and len(b.color.rgb) >= 6:
                    color = "#" + b.color.rgb[-6:]
            except Exception:
                pass
            width = self.PEN_HAIR if b.style == "hair" else self.PEN_THIN
            return QPen(QColor(color), width)

        def draw(side, x1, y1, x2, y2):
            p = pen(side)
            if p is None:
                return
            painter.setPen(p)
            painter.drawLine(x1, y1, x2, y2)

        left_c = mr[2] if mr else c
        right_c = mr[3] if mr else c
        top_r = mr[0] if mr else r
        bot_r = mr[1] if mr else r
        draw("left", rect.left(), rect.top(), rect.left(), rect.bottom())
        if c == right_c:
            draw("right", rect.right(), rect.top(), rect.right(), rect.bottom())
        draw("top", rect.left(), rect.top(), rect.right(), rect.top())
        if r == bot_r:
            draw("bottom", rect.left(), rect.bottom(), rect.right(), rect.bottom())

    def _text(self, painter, r, c, cell, text):
        m = self.m
        rect = m.cell_rect(r, c)
        inset = 1.5
        box = rect.adjusted(inset, inset, -inset, -inset)
        font = self.qfont(cell)
        al = cell.alignment
        painter.setFont(font)
        painter.setPen(QColor("#111111"))
        fm = QFontMetrics(font)

        lines = []
        if al.wrap_text:
            for para in text.split("\n"):
                cur = ""
                for word in para.split(" "):
                    cand = (cur + " " + word).strip()
                    if fm.horizontalAdvance(cand) <= box.width() or not cur:
                        cur = cand
                    else:
                        lines.append(cur)
                        cur = word
                lines.append(cur)
        else:
            lines = text.split("\n")

        rot = al.textRotation or 0
        if rot == 90:
            painter.save()
            painter.translate(rect.center())
            painter.rotate(-90)
            w, h = box.height(), box.width()
            total = len(lines) * fm.height()
            y = -total / 2
            for ln in lines:
                tw = fm.horizontalAdvance(ln)
                x = -tw / 2 if al.horizontal in (None, "center") else (-w / 2 if al.horizontal == "left" else w / 2 - tw)
                painter.drawText(int(x), int(y + fm.ascent()), ln)
                y += fm.height()
            painter.restore()
            return

        total = len(lines) * fm.height()
        y = box.top()
        if al.vertical == "center":
            y = box.top() + (box.height() - total) / 2
        elif al.vertical == "bottom":
            y = box.bottom() - total
        for ln in lines:
            tw = fm.horizontalAdvance(ln)
            if al.horizontal == "center":
                x = box.left() + (box.width() - tw) / 2
            elif al.horizontal == "right":
                x = box.right() - tw
            else:
                x = box.left()
            painter.drawText(int(x), int(y + fm.ascent()), ln)
            y += fm.height()


# ──────────────────────────────────────────────────────────────────
# Вывод: принтер / PDF / картинка
# ──────────────────────────────────────────────────────────────────

def _apply_pages(pdevice, model, orientation, paper_size):
    from PySide6.QtGui import QPageLayout, QPageSize
    size = QPageSize(QPageSize.A3 if paper_size == "A3" else QPageSize.A4)
    orient = QPageLayout.Landscape if orientation == "landscape" else QPageLayout.Portrait
    pdevice.setPageSize(size)
    pdevice.setPageOrientation(orient)
    ml, mr, mt, mb = model.margins_mm
    pdevice.setPageMargins(QMarginsF(ml, mt, mr, mb), QPageLayout.Millimeter)


def render_to_device(model, pdevice, page_from=None, page_to=None):
    """Рисует лист на QPrinter/QPdfWriter. Возвращает число напечатанных страниц."""
    renderer = SheetRenderer(model)
    painter = QPainter(pdevice)
    try:
        # печатная область в листовых px с учётом масштаба шаблона
        pr = pdevice.pageLayout().paintRectPixels(pdevice.resolution())
        k_base = pdevice.resolution() / 96.0 * model.scale
        avail_w_px = pr.width() / k_base
        if model.width > avail_w_px:          # чуть не влезает — прижимаем ширину
            k_base *= avail_w_px / model.width
        page_h_px = pr.height() / k_base
        pages = renderer.page_rows(page_h_px)
        pages = pages[int(page_from) - 1: int(page_to)] if (page_from and page_to) else pages
        for i, rows in enumerate(pages):
            if i:
                pdevice.newPage()
            painter.save()
            painter.setTransform(QTransform())
            painter.scale(k_base, k_base)
            renderer.paint(painter, rows)
            painter.restore()
        return len(pages)
    finally:
        painter.end()


def sheet_to_pdf(ws, out_path, orientation="landscape", paper_size="A4"):
    from PySide6.QtGui import QPdfWriter
    model = SheetModel(ws)
    pdf = QPdfWriter(out_path)
    pdf.setResolution(96)
    _apply_pages(pdf, model, orientation, paper_size)
    return render_to_device(model, pdf)


def sheet_to_image(ws, out_path, dpi=96):
    """Превью первого фрагмента листа (для тестов)."""
    model = SheetModel(ws)
    renderer = SheetRenderer(model, font_hook=_HOOK or _system_font_hook)
    img = QImage(model.width, min(model.height, 5000), QImage.Format_RGB32)
    img.fill(QColor("#ffffff"))
    p = QPainter(img)
    p.setTransform(QTransform())
    renderer.paint(p, list(range(model.r0, model.r1 + 1)))
    p.end()
    img.save(out_path)
    return model


# подмена шрифтов для песочницы (нет PT Astra Serif); в Windows не активна
def _sandbox_hook(name):
    try:
        import glob
        ttf = glob.glob("/usr/local/lib/python3.11/dist-packages/matplotlib/"
                        "mpl-data/fonts/ttf/DejaVuSerif*.ttf")
        if not ttf:
            return name if name != "PT Astra Serif" else None
        from PySide6.QtGui import QGuiApplication, QFontDatabase
        if not QGuiApplication.instance():
            return None
        if not getattr(_sandbox_hook, "_loaded", False):
            from PySide6.QtGui import QFontDatabase as FD
            for f in ttf:
                FD.addApplicationFont(f)
            sans = glob.glob("/usr/local/lib/python3.11/dist-packages/matplotlib/"
                             "mpl-data/fonts/ttf/DejaVuSans.ttf")
            for f in sans:
                FD.addApplicationFont(f)
            _sandbox_hook._loaded = True
        return "DejaVu Serif" if "Serif" in name or "Astra" in name else "DejaVu Sans"
    except Exception:
        return None


def _system_font_hook(name):
    """Обычный режим: имя проходит, если шрифт есть в системе."""
    try:
        from PySide6.QtGui import QFontDatabase
        return name if name in QFontDatabase.families() else None
    except Exception:
        return name


_HOOK = _sandbox_hook if os.environ.get("OVERTIMETAB_SANDBOX_FONTS") else None


def print_sheet_to_printer(ws, printer_name, copies, page_from, page_to,
                           orientation, paper_size, collate):
    printer = QPrinter(QPrinter.HighResolution)
    if printer_name:
        printer.setPrinterName(printer_name)
    printer.setCopyCount(int(copies or 1))
    printer.setCollate(bool(collate))
    model = SheetModel(ws)
    _apply_pages(printer, model, orientation, paper_size)
    pf, pt = None, None
    try:
        if str(page_from).strip() and str(page_to).strip():
            pf, pt = int(page_from), int(page_to)
            printer.setFromTo(pf, pt)
    except ValueError:
        pass
    if not printer.isValid():
        raise RuntimeError("Принтер «%s» недоступен" % (printer_name or "по умолчанию"))
    return render_to_device(model, printer, pf, pt)


# ──────────────────────────────────────────────────────────────────
# Высокий уровень: база → заполненный бланк → печать
# ──────────────────────────────────────────────────────────────────

def print_report(db_path, year, month, template_path, printer_name, copies,
                 page_from, page_to, orientation, paper_size, collate,
                 sheet_name="Лист1") -> int:
    """Формирует бланк и печатает его без Excel. Возвращает число страниц."""
    from database import DB
    from export import TemplateExporter

    temp_db = DB(db_path)
    try:
        temp_dir = Path(tempfile.gettempdir())
        temp_xlsx = temp_dir / ("overtimetab_qtprint_%s.xlsx" % uuid.uuid4().hex[:6])
        TemplateExporter.export(db=temp_db, year=year, month=month,
                                template_path=template_path, out_path=str(temp_xlsx),
                                sheet_name=sheet_name)
    finally:
        temp_db.close()

    try:
        wb = openpyxl.load_workbook(temp_xlsx)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        n = print_sheet_to_printer(ws, printer_name, copies, page_from, page_to,
                                   orientation, paper_size, collate)
        return n
    finally:
        try:
            os.remove(str(temp_xlsx))
        except OSError:
            pass
