#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ВЕБ-ВЕРСИЯ OVERTIMETAB (этап переезда) — вертикальный срез для решения о миграции на HTML.

Работает на НАСТОЯЩЕМ ядре программы: database.py, logic.py и utils.py
из корня репозитория. Ни одного выдуманного числа: календарь, дежурства,
остатки и «всего дней» считаются теми же функциями, что и в QML-версии.

Запуск:  python3 webapp/server.py   →  http://localhost:8081
Только стандартная библиотека + модули самого приложения.
"""
import json
import os
import re
import sys
import calendar as cal_lib
from datetime import date, datetime, timedelta
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(ROOT))          # доступ к ядру приложения
from database import DB                             # noqa: E402
from logic import (compute_month_summary,           # noqa: E402
                   build_shifted_weekend_checker,
                   resolve_is_working)
from utils import (d_iso, dt_parse, subtract_intervals, intersect,  # noqa: E402
                   month_bounds_dt, fmt_minutes_ru_words, parse_hhmm)

DEMO_DB = os.path.join(ROOT, "demo.db")
PORT = 8081

# Реальная база — только по явному указанию (флаг --db или переменная
# окружения OVERTIMETAB_DB). Без него работает демо-база: постепенный
# переезд не должен рисковать данными пользователей.
REAL_DB = None

# Сколько правок сделала ЭТА сессия сервера: отмена не должна трогать
# снапшоты, оставшиеся в чужой базе от прошлых запусков программы.
_session_undo = 0


def bump_undo():
    global _session_undo
    _session_undo += 1

# ════════════════════════════════════════════════════════════════════
# ОФОРМЛЕНИЕ ОСТАТКОВ — порт из Main.py (там они рядом с PySide6,
# поэтому импортировать Main.py напрямую нельзя). Логика идентична.
# ════════════════════════════════════════════════════════════════════
GRAY = "#9E9E9E"


def fmt_dual(val, prev_val, is_days=False):
    if prev_val == 0:
        return (f"{val} д." if is_days else fmt_minutes_ru_words(val))
    if is_days:
        return f"{val} <span class='prev'>({prev_val})</span> д."
    av, pv = abs(int(val)), abs(int(prev_val))
    if av % 60 == 0 and pv % 60 == 0:
        sv = "-" if val < 0 else ""
        sp = "-" if prev_val < 0 else ""
        return f"{sv}{av // 60} <span class='prev'>({sp}{pv // 60})</span> ч."
    main = fmt_minutes_ru_words(val)
    prev_txt = fmt_minutes_ru_words(prev_val)
    return f"{main} <span class='prev'>({prev_txt})</span>"


def fmt_comp(real_val, prev_val, is_days=False):
    if real_val == 0 and prev_val == 0:
        return "—"
    return fmt_dual(real_val, prev_val, is_days)


def total_overtime_days(s):
    night = int(s.get("end_hours", 0)) + int(s.get("prev_h_end", 0))
    extra = int(s.get("end_overtime", 0)) + int(s.get("prev_o_end", 0))
    if extra < 0:
        extra = 0
    return (night + extra) // (8 * 60) + int(s.get("end_days", 0)) + int(s.get("prev_d_end", 0))


def plain(html):
    """HTML-строка → чистый текст (для одометра)."""
    return re.sub(r"<[^>]+>", "", html).strip()


MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
             "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

# ════════════════════════════════════════════════════════════════════
# ДЕМО-БАЗА: создаётся один раз через САМИ функции программы.
# Персонажи с разными судьбами — чтобы все состояния интерфейса
# были видны: накопитель, тратящий, новичок, отпуск, больничный, нули.
# ════════════════════════════════════════════════════════════════════

def seed(db: DB):
    g_sut    = db.add_group("Суточный наряд", is_shift=True)
    g_night  = db.add_group("Ночная смена", is_shift=True)
    g_day    = db.add_group("Пятидневка", is_shift=False)

    def emp(last, first, middle, position, start, group, **kw):
        return db.add_employee(last, first, middle, kw.get("rank", ""),
                               position, start, 0, 0, 0, 0, 0, 0, group)

    # start_month = месяц, с которого сеются дежурства: норма в движке
    # считается с приёма, расхождение дало бы фиктивные минусы.
    ivan   = emp("Иванов", "Иван", "Иванович", "инженер", "2025-01", g_sut)
    petrov = emp("Петров", "Пётр", "Петрович", "техник", "2025-06", g_sut)
    sid    = emp("Сидорова", "Анна", "Александровна", "диспетчер", "2025-01", g_day)
    kuz    = emp("Кузнецов", "Денис", "Сергеевич", "инженер", "2025-06", g_sut)
    smir   = emp("Смирнов", "Алексей", "Викторович", "техник", "2025-03", g_night)
    volk   = emp("Волкова", "Мария", "Игоревна", "диспетчер", "2025-11", g_day)
    novik  = emp("Новиков", "Игорь", "Павлович", "инженер", "2026-06", g_sut)
    fed    = emp("Фёдорова", "Ольга", "Николаевна", "техник", "2025-05", g_night)
    dmit   = emp("Дмитриев", "Сергей", "Сергеевич", "начальник караула", "2025-01", g_day)

    # ── производственный календарь 2026 (праздники) ──────────────
    holidays_2026 = (
        [date(2026, 1, d) for d in range(1, 9)] +        # новогодние
        [date(2026, 2, 23), date(2026, 3, 8), date(2026, 5, 1),
         date(2026, 5, 9), date(2026, 6, 12), date(2026, 11, 4)]
    )
    for d0 in holidays_2026:
        db.set_calendar_day_type(d0, "holiday")
    # перенесённые выходные 2026 (нерабочие, но не праздники)
    for d0 in (date(2026, 1, 9), date(2026, 3, 9), date(2026, 5, 11)):
        db.set_calendar_day_type(d0, "weekend")
    # 2025 — грубо, только чтобы остатки прошлого года считались честно
    holidays_2025 = (
        [date(2025, 1, d) for d in range(1, 9)] +
        [date(2025, 2, 23), date(2025, 3, 8), date(2025, 5, 1), date(2025, 5, 8),
         date(2025, 5, 9), date(2025, 6, 12), date(2025, 11, 3), date(2025, 11, 4)]
    )
    for d0 in holidays_2025:
        db.set_calendar_day_type(d0, "holiday")

    # ── дежурства ────────────────────────────────────────────────
    def every_n(eid, y1, m1, y2, m2, n, start_h, hours, is_shift, comment="", skip_months=()):
        """Дежурство каждые n дней с y1-m1 по y2-m2 (включительно)."""
        d0, d1 = date(y1, m1, 1), date(y2, m2, 1)
        d1 = (d1.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        d = d0
        ordinal = d.toordinal()
        while d <= d1:
            if d.month not in skip_months:
                s = datetime(d.year, d.month, d.day, start_h, 0)
                db.add_duty(eid, s, s + timedelta(hours=hours), comment, is_shift)
            ordinal += n
            d = date.fromordinal(ordinal)

    today = date.today()
    end_y, end_m = today.year, today.month

    every_n(ivan,  2025, 1, 2025, 12, 3, 8, 24, True)          # сутки через двое, 2025
    every_n(ivan,  2026, 1, end_y, end_m, 3, 8, 24, True)      # 2026
    every_n(petrov, 2025, 6, 2025, 12, 3, 8, 24, True)         # 2025 плотно
    every_n(petrov, 2026, 1, end_y, end_m, 4, 8, 24, True)     # 2026 — сутки через трое
    every_n(kuz,   2025, 6, 2025, 12, 3, 9, 24, True)
    every_n(kuz,   2026, 1, end_y, end_m, 3, 9, 24, True)      # 2026 — интенсивно копит
    every_n(novik, 2026, 6, end_y, end_m, 4, 8, 24, True)      # новичок с июня
    every_n(smir,  2025, 3, 2025, 12, 2, 20, 12, True)         # ночные 2/2
    every_n(smir,  2026, 1, end_y, end_m, 2, 20, 12, True)
    every_n(fed,   2025, 5, 2025, 12, 2, 20, 12, True)
    every_n(fed,   2026, 1, end_y, end_m, 2, 20, 12, True)

    # у Иванова одно дежурство в августе 2026 — с обеденным перерывом
    aug = [r for r in db.list_duties_for_period(
        ivan, datetime(2026, 8, 1), datetime(2026, 9, 1))
        if dt_parse(r["start_dt"]).day == 15]
    if aug:
        did = int(aug[0]["id"])
        s = dt_parse(aug[0]["start_dt"])
        db.replace_duty_breaks(did, [(s + timedelta(hours=4), s + timedelta(hours=5))])

    # пятидневка: редкие дежурства по выходным = переработка (is_shift=False)
    for m in range(1, end_m + 1):
        for day in (10, 24):                                    # субботы-воскресенья условно
            try:
                d0 = date(2026, m, day)
            except ValueError:
                continue
            if d0.weekday() >= 5 and d0 <= today:
                db.add_duty(sid, datetime(2026, m, day, 8, 0),
                            datetime(2026, m, day, 16, 0), "", False)
                if d0.day == 24 and d0.month in (2, 6, 9):
                    db.add_duty(volk, datetime(2026, m, day, 8, 0),
                                datetime(2026, m, day, 13, 0), "", False)

    # ── статусы ──────────────────────────────────────────────────
    for d in range(6, 11):                                     # отпуск Смирнова, июль
        db.set_day_status(smir, date(2026, 7, d), "О")
    # дежурства, попавшие на отпуск, убираем — данные должны быть честными
    with db.conn:
        db.conn.execute(
            "DELETE FROM duty WHERE employee_id=? AND date(start_dt) BETWEEN '2026-07-06' AND '2026-07-10'",
            (smir,))
    for d in range(10, 15):                                     # больничный Волковой
        db.set_day_status(volk, date(2026, 8, d), "Б")

    # ── компенсации ──────────────────────────────────────────────
    # Иванов регулярно «душит» накопления часовыми отгулами
    for mn in (2, 4, 6, 8):
        db.add_compensation_hours_dayoff(ivan, date(2026, mn, 12), 1440, "Приказ")
    db.add_compensation_days_dayoff(kuz, [date(2026, 3, 16), date(2026, 3, 17),
                                          date(2026, 3, 18)], "Приказ №41")
    db.add_compensation_hours_dayoff(kuz, date(2026, 6, 8), 480, "Приказ №77")
    db.add_compensation_hours_dayoff(fed, date(2026, 4, 10), 480, "Приказ №52")
    db.add_compensation_hours_dayoff(fed, date(2026, 8, 7), 480, "Приказ №103")

    db.conn.commit()


# ════════════════════════════════════════════════════════════════════
# ПОСТРОЕНИЕ ДАННЫХ (порты из Main.py — та же логика, что в QML)
# ════════════════════════════════════════════════════════════════════

_mini_cache = {}    # (emp_id, year) → [12 значений «всего дней»]


def mini_year(db, emp_id, year):
    key = (emp_id, year)
    if key not in _mini_cache:
        out = []
        for m in range(1, 13):
            s = compute_month_summary(db, emp_id, year, m)
            out.append(int(total_overtime_days(s)))
        _mini_cache[key] = out
    return _mini_cache[key]


def build_month_grid(db, emp_id, year, month):
    weeks = cal_lib.Calendar(firstweekday=0).monthdatescalendar(year, month)
    grid_start, grid_end = weeks[0][0], weeks[-1][-1]
    gs, ge = d_iso(grid_start), d_iso(grid_end)

    work_map = db.get_calendar_month(gs, ge)
    holidays = db.get_holidays_month(gs, ge)
    overrides = db.get_calendar_overrides(gs, ge)
    pre_holidays = db.get_pre_holidays_month(gs, ge)
    shifted = build_shifted_weekend_checker(db, emp_id) if emp_id else None

    statuses = db.get_statuses_for_period(emp_id, gs, ge) if emp_id else {}
    comp_set = set()
    if emp_id:
        for c in db.list_compensations_for_period(emp_id, gs, ge):
            if c["unit"] in ("hours", "overtime"):
                dd = c["order_date"] if c["order_date"] else c["event_date"]
                if dd:
                    comp_set.add(dd)
            else:
                for cd in db.get_comp_dates(int(c["id"])):
                    comp_set.add(cd)

    duty_map = {}
    if emp_id:
        s_dt = datetime.combine(grid_start, datetime.min.time())
        e_dt = datetime.combine(grid_end + timedelta(days=1), datetime.min.time())
        for d in db.list_duties_for_period(emp_id, s_dt, e_dt):
            did = int(d["id"])
            is_shift = bool(int(d["is_shift"] or 0))
            s0 = max(dt_parse(d["start_dt"]), s_dt)
            e0 = min(dt_parse(d["end_dt"]), e_dt)
            if s0 >= e0:
                continue
            parts = subtract_intervals((s0, e0), [])
            for s, e in parts:
                cur = s.date()
                last = (e - timedelta(seconds=1)).date() if e > s else s.date()
                while cur <= last:
                    ds = datetime.combine(cur, datetime.min.time())
                    de = ds + timedelta(days=1)
                    inter = intersect(s, e, ds, de)
                    if inter:
                        t_str = f"{inter[0].strftime('%H:%M')}-{inter[1].strftime('%H:%M')}"
                        duty_map.setdefault(d_iso(cur), []).append(
                            {"id": did, "text": t_str, "is_shift": is_shift})
                    cur += timedelta(days=1)

    days = []
    for wk in weeks:
        for d in wk:
            d_str = d_iso(d)
            is_working = resolve_is_working(
                d, bool(shifted and shifted(d)), work_map, holidays, overrides)
            is_holiday = d in holidays
            days.append({
                "date": d_str,
                "n": d.day,
                "in_month": d.month == month,
                "is_weekend": (not is_working) and (not is_holiday),
                "is_holiday": is_holiday,
                "is_pre_holiday": d in pre_holidays,
                "status": statuses.get(d, "") if isinstance(statuses.get(d, ""), str) else "",
                "has_comp": d_str in comp_set,
                "duties": duty_map.get(d_str, []),
            })
    return days


def build_summary(db, emp_id, year, month):
    s = compute_month_summary(db, emp_id, year, month)

    def card(title, unit_key, is_days):
        end_html = fmt_dual(s.get(f"end_{unit_key}", 0), s.get(f"prev_{unit_key[0]}_end", 0), is_days)
        return {
            "title": title,
            "start": fmt_dual(s.get(f"start_{unit_key}", 0), s.get(f"prev_{unit_key[0]}_start", 0), is_days),
            "acc": (f"{s.get('acc_' + unit_key, 0)} д." if is_days
                    else fmt_minutes_ru_words(s.get("acc_" + unit_key, 0))),
            "comp": fmt_comp(s.get(f"comp_{unit_key[0]}_real", 0), s.get(f"comp_{unit_key[0]}_prev", 0), is_days),
            "end_html": end_html,
            "end_plain": plain(end_html),
            "end_neg": s.get(f"end_{unit_key}", 0) < 0,
        }

    return {
        "total_days": int(total_overtime_days(s)),
        "norm": fmt_minutes_ru_words(s.get("norm_minutes", 0)),
        "shift": fmt_minutes_ru_words(s.get("shift_minutes", 0)),
        "night": fmt_minutes_ru_words(s.get("shift_night_minutes", 0)),
        "is_shift_month": s.get("is_shift", False) is True,
        "cards": [
            card("Ночные (ДВО)", "hours", False),
            card("Сверх нормы", "overtime", False),
            card("Дни (ДДО)", "days", True),
        ],
    }


def build_year(db, emp_id, year):
    """Панорама года: по месяцам — матрица дней с минутами дежурств."""
    s_dt = datetime(year, 1, 1)
    e_dt = datetime(year + 1, 1, 1)
    per_day = {}
    for d in db.list_duties_for_period(emp_id, s_dt, e_dt):
        s0, e0 = dt_parse(d["start_dt"]), dt_parse(d["end_dt"])
        cur = s0.date()
        while cur <= (e0 - timedelta(seconds=1)).date():
            ds = datetime.combine(cur, datetime.min.time())
            de = ds + timedelta(days=1)
            inter = intersect(s0, e0, ds, de)
            if inter:
                mins = int((inter[1] - inter[0]).total_seconds() // 60)
                per_day[cur.isoformat()] = per_day.get(cur.isoformat(), 0) + mins
            cur += timedelta(days=1)
    holidays = db.get_holidays_month(d_iso(date(year, 1, 1)), d_iso(date(year, 12, 31)))
    months = []
    for m in range(1, 13):
        weeks = cal_lib.Calendar(firstweekday=0).monthdatescalendar(year, m)
        s = compute_month_summary(db, emp_id, year, m)
        months.append({
            "m": m,
            "weeks": [[d.isoformat() for d in wk] for wk in weeks],
            "total": int(total_overtime_days(s)),
            "neg": s.get("end_overtime", 0) < 0 or s.get("end_days", 0) < 0,
        })
    return {"months": months, "per_day": per_day,
            "holidays": [d.isoformat() for d in holidays]}


def fio(e):
    initials = ((e["first_name"] or "")[:1] + ".") if e["first_name"] else ""
    initials += ((e["middle_name"] or "")[:1] + ".") if e["middle_name"] else ""
    return f"{e['last_name']} {initials}".strip()


# ════════════════════════════════════════════════════════════════════
# ДЕНЬ: данные инспектора (дежурства с перерывами, компенсации, статус)
# ════════════════════════════════════════════════════════════════════

def build_day(db, emp_id, d: date):
    day_start = datetime.combine(d, datetime.min.time())
    day_end = day_start + timedelta(days=1)
    duties = []
    for r in db.list_duties_for_period(emp_id, day_start, day_end):
        s0, e0 = dt_parse(r["start_dt"]), dt_parse(r["end_dt"])
        inter = intersect(s0, e0, day_start, day_end)
        if not inter:
            continue
        duties.append({
            "id": int(r["id"]),
            "is_shift": bool(int(r["is_shift"] or 0)),
            "comment": r["comment"] or "",
            "start": r["start_dt"][:16],
            "end": r["end_dt"][:16],
            "slice": f"{inter[0].strftime('%H:%M')}-{inter[1].strftime('%H:%M')}",
            "multi": s0.date() != (e0 - timedelta(seconds=1)).date(),
        })
    breaks_map = db.breaks_for_duty_ids([x["id"] for x in duties])
    for x in duties:
        x["breaks"] = [
            f"{dt_parse(b[0]).strftime('%H:%M')}-{dt_parse(b[1]).strftime('%H:%M')}"
            for b in breaks_map.get(x["id"], [])]

    comps = []
    for c in db.list_compensations_for_period(emp_id, d_iso(d), d_iso(d)):
        if c["unit"] in ("hours", "overtime"):
            dd = c["order_date"] if c["order_date"] else c["event_date"]
            if dd == d_iso(d):
                h = int(c["amount_minutes"] or 0) / 60
                comps.append({"kind": "hours",
                              "text": f"{'сверх нормы' if c['unit'] == 'overtime' else 'ДВО'} · {h:g} ч."})
        else:
            if d_iso(d) in db.get_comp_dates(int(c["id"])):
                comps.append({"kind": "days", "text": f"отгул · {c['amount_days']} дн."})

    gs, ge = d_iso(d), d_iso(d)
    work_map = db.get_calendar_month(gs, ge)
    holidays = db.get_holidays_month(gs, ge)
    overrides = db.get_calendar_overrides(gs, ge)
    pre = db.get_pre_holidays_month(gs, ge)
    shifted = build_shifted_weekend_checker(db, emp_id)
    is_working = resolve_is_working(d, bool(shifted and shifted(d)),
                                    work_map, holidays, overrides)
    statuses = db.get_statuses_for_period(emp_id, gs, ge)

    return {
        "date": d_iso(d),
        "is_working": is_working,
        "is_holiday": d in holidays,
        "is_pre_holiday": d in pre,
        "status": statuses.get(d, ""),
        "duties": duties,
        "comps": comps,
    }


def month_payload(db, emp_id, year, month):
    """Полный ответ «месяц» — им же отвечают правки, чтобы интерфейс
    обновлялся за один круг."""
    e = db.get_employee(emp_id)
    s = compute_month_summary(db, emp_id, year, month)
    norm = s.get("norm_minutes", 0)
    return {
        "emp": {"id": emp_id, "fio": fio(e), "position": e["position"] or ""},
        "year": year, "month": month,
        "period": f"{MONTHS_RU[month - 1]} {year}",
        "days": build_month_grid(db, emp_id, year, month),
        "summary": build_summary(db, emp_id, year, month),
        "mini": mini_year(db, emp_id, year),
        "ratio": round(max(0, s.get("shift_minutes", 0)) / norm, 3) if norm else 0,
    }


# ════════════════════════════════════════════════════════════════════
# HTTP
# ════════════════════════════════════════════════════════════════════

def open_db():
    if REAL_DB:
        print("Открываю реальную базу:", REAL_DB)
        return DB(REAL_DB)
    fresh = not os.path.exists(DEMO_DB)
    db = DB(DEMO_DB)
    if fresh:
        seed(db)
        print("Демо-база создана и наполнена:", DEMO_DB)
    return db


class Handler(SimpleHTTPRequestHandler):
    db: DB = None

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=os.path.join(ROOT, "static"), **kw)

    def log_message(self, fmt, *args):
        pass  # тише в консоли

    def json_out(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def body_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        try:
            if u.path == "/api/bootstrap":
                self.api_bootstrap(q)
            elif u.path == "/api/month":
                self.api_month(q)
            elif u.path == "/api/year":
                self.api_year(q)
            elif u.path == "/api/day":
                self.api_day(q)
            else:
                super().do_GET()
        except Exception as ex:  # noqa: BLE001
            self.json_out({"error": f"{type(ex).__name__}: {ex}"}, 500)
            import traceback
            traceback.print_exc()

    def do_POST(self):
        u = urlparse(self.path)
        try:
            if u.path == "/api/duty/add":
                self.api_duty_add()
            elif u.path == "/api/duty/delete":
                self.api_duty_delete()
            elif u.path == "/api/day/set-status":
                self.api_set_status()
            elif u.path == "/api/undo":
                self.api_undo()
            else:
                self.json_out({"error": "неизвестный запрос"}, 404)
        except Exception as ex:  # noqa: BLE001
            self.json_out({"error": f"{type(ex).__name__}: {ex}"}, 500)
            import traceback
            traceback.print_exc()

    def after_write(self, emp_id, year, month, message, kind="success"):
        """После правки: сброс кэшей, свежие данные месяца одним ответом."""
        _mini_cache.clear()
        self.json_out({"ok": True, "message": message, "kind": kind,
                       "month": month_payload(self.db, emp_id, year, month)})

    # ── /api/bootstrap?year=2026 ────────────────────────────────
    def api_bootstrap(self, q):
        db = self.db
        year = int(q.get("year", [date.today().year])[0])
        month = date.today().month
        groups = [{"id": int(g["id"]), "name": g["name"]}
                  for g in db.list_groups()]
        employees = []
        for e in db.list_employees_for_month(year, month, True):
            eid = int(e["id"])
            s = compute_month_summary(db, eid, year, month)
            norm = s.get("norm_minutes", 0)
            employees.append({
                "id": eid, "fio": fio(e),
                "position": e["position"] or "",
                "group_id": int(e["group_id"]) if e["group_id"] is not None else None,
                "start_month": e["start_month"],
                "ratio": round(max(0, s.get("shift_minutes", 0)) / norm, 3) if norm else 0,
                "mini": mini_year(db, eid, year),
            })
        self.json_out({
            "today": {"y": date.today().year, "m": date.today().month,
                      "d": date.today().day},
            "year": year,
            "groups": groups,
            "employees": employees,
        })

    # ── /api/month?emp=1&year=2026&month=9 ──────────────────────
    def api_month(self, q):
        db = self.db
        emp = int(q.get("emp", [0])[0])
        year = int(q.get("year", [date.today().year])[0])
        month = int(q.get("month", [date.today().month])[0])
        self.json_out(month_payload(db, emp, year, month))

    # ── /api/day?emp=1&date=2026-09-23 ──────────────────────────
    def api_day(self, q):
        db = self.db
        emp = int(q.get("emp", [0])[0])
        d = date.fromisoformat(q.get("date", [date.today().isoformat()])[0])
        self.json_out(build_day(db, emp, d))

    # ── POST /api/duty/add ─────────────────────────────────────
    def api_duty_add(self):
        db = self.db
        b = self.body_json()
        emp = int(b["emp"])
        d0 = date.fromisoformat(b["date"])
        try:
            s_t = parse_hhmm(b.get("start", ""))
            e_t = parse_hhmm(b.get("end", ""))
        except Exception:
            self.json_out({"ok": False, "message": "Неверное время"}, 400)
            return
        start_dt = datetime.combine(d0, s_t)
        end_dt = datetime.combine(d0, e_t)
        if end_dt <= start_dt:
            end_dt += timedelta(days=1)          # дежурство через полночь

        overlaps = db.find_overlapping_duties(emp, start_dt, end_dt)
        if overlaps:
            self.json_out({"ok": False,
                           "message": "Ошибка: пересечение с другим дежурством"},
                          409)
            return
        with db.transaction():
            db.add_duty(emp, start_dt, end_dt, b.get("comment", ""),
                        bool(b.get("is_shift", False)))
        bump_undo()
        self.after_write(emp, d0.year, d0.month, "Дежурство сохранено")

    # ── POST /api/duty/delete {id, emp, year, month} ───────────
    def api_duty_delete(self):
        db = self.db
        b = self.body_json()
        with db.transaction():
            db.delete_duty(int(b["id"]))
        bump_undo()
        self.after_write(int(b["emp"]), int(b.get("year", date.today().year)),
                         int(b.get("month", date.today().month)),
                         "Дежурство удалено")

    # ── POST /api/day/set-status {emp, date, status} ───────────
    def api_set_status(self):
        db = self.db
        b = self.body_json()
        emp = int(b["emp"])
        d0 = date.fromisoformat(b["date"])
        status = b.get("status", "")
        if status not in ("", "К", "Б", "О"):
            self.json_out({"ok": False, "message": "Неизвестный статус"}, 400)
            return
        with db.transaction():
            db.set_day_status(emp, d0, status)
        bump_undo()
        msg = "Статус снят" if not status else f"Статус: {status}"
        self.after_write(emp, d0.year, d0.month, msg)

    # ── POST /api/undo ─────────────────────────────────────────
    def api_undo(self):
        global _session_undo
        db = self.db
        if _session_undo <= 0:
            self.json_out({"ok": False, "message": "Нечего отменять",
                           "kind": "error"})
            return
        if not db.undo():
            _session_undo = 0
            self.json_out({"ok": False, "message": "Нечего отменять",
                           "kind": "error"})
            return
        _session_undo -= 1
        _mini_cache.clear()
        # отменённое могло быть в любом месяце — интерфейс перечитает всё
        self.json_out({"ok": True, "message": "Действие отменено",
                       "reload": True})

    # ── /api/year?emp=1&year=2026 ───────────────────────────────
    def api_year(self, q):
        db = self.db
        emp = int(q.get("emp", [0])[0])
        year = int(q.get("year", [date.today().year])[0])
        e = db.get_employee(emp)
        self.json_out({
            "emp": {"id": emp, "fio": fio(e)},
            "year": year,
            "panorama": build_year(db, emp, year),
            "mini": mini_year(db, emp, year),
        })


def main():
    global REAL_DB
    args = sys.argv[1:]
    if "--db" in args:
        REAL_DB = args[args.index("--db") + 1]
    elif os.environ.get("OVERTIMETAB_DB"):
        REAL_DB = os.environ["OVERTIMETAB_DB"]
    if REAL_DB and not os.path.exists(REAL_DB):
        print("База не найдена:", REAL_DB)
        sys.exit(1)
    Handler.db = open_db()
    httpd = HTTPServer(("0.0.0.0", PORT), Handler)
    mode = f"РЕАЛЬНАЯ база: {REAL_DB}" if REAL_DB else "демо-база"
    print(f"Веб-версия OVERTIMETAB: http://0.0.0.0:{PORT}  ({mode})")
    httpd.serve_forever()


if __name__ == "__main__":
    main()
