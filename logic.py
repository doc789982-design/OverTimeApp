from datetime import date, datetime, time, timedelta
from typing import Optional
from utils import d_iso, d_parse, dt_parse, dt_iso, subtract_intervals, merge_intervals, intersect, month_bounds_dt

def safe_get_hire_date(start_str):
    """Безопасно достает год и месяц приема. Если в базе пусто - берет 2000-01"""
    if not start_str or "-" not in str(start_str):
        return 2000, 1
    try:
        y, m = start_str.split("-")
        return int(y), int(m)
    except:
        return 2000, 1

def extract_night_intervals(s: datetime, e: datetime) -> list[tuple[datetime, datetime]]:
    out = []
    day = s.date() - timedelta(days=1)
    last_day = e.date()
    while day <= last_day:
        w0 = datetime.combine(day, time(22, 0))
        w1 = datetime.combine(day + timedelta(days=1), time(6, 0))
        inter = intersect(s, e, w0, w1)
        if inter: out.append(inter)
        day += timedelta(days=1)
    return out

def build_shift_checker(db, employee_id: int):
    groups = db.list_groups()
    shift_map = {g["id"]: bool(int(g["is_shift"])) for g in groups}
    transfers = db.conn.execute("SELECT transfer_date, group_id FROM employee_transfer WHERE employee_id=? ORDER BY transfer_date DESC", (employee_id,)).fetchall()
    history = [(d_parse(t["transfer_date"]), t["group_id"]) for t in transfers]
    emp = db.get_employee(employee_id)
    current_gid = emp["group_id"]
    def check(target_date: Optional[date] = None) -> bool:
        if target_date is None: return shift_map.get(current_gid, False)
        for t_date, gid in history:
            if target_date >= t_date: return shift_map.get(gid, False)
        return shift_map.get(current_gid, False)
    return check

def is_employee_shift(db, employee_id: int, target_date: Optional[date] = None) -> bool:
    return build_shift_checker(db, employee_id)(target_date)

def default_is_working(d: date, shifted: bool = False) -> bool:
    """Шаблон недели, если дня нет в calendar_day.
    Обычный: Пн–Пт. Смещённый: Вт–Сб (пн и вс выходные)."""
    wd = d.weekday()
    if shifted:
        return wd not in (0, 6)
    return wd < 5

def resolve_is_working(
    d: date,
    shifted: bool,
    work_map: dict,
    holidays_set: Optional[set] = None,
) -> bool:
    """Рабочий ли день для этого сотрудника.

    Праздники общие. Явное исключение из обычной пятидневки
    (перенос, «сделать рабочим/выходным») тоже общее.
    Строки calendar_day, которые просто повторяют пн–пт / сб–вс,
    шаблон группы не затирают — иначе смещённые выходные пропадают
    в уже заполненном календаре.
    """
    holidays_set = holidays_set or set()
    stored = work_map.get(d)
    if d in holidays_set:
        return False if stored is None else bool(stored)
    if stored is None:
        return default_is_working(d, shifted)
    if not shifted:
        return bool(stored)
    if bool(stored) != default_is_working(d, False):
        return bool(stored)
    return default_is_working(d, True)

def _row_flag(row, key: str) -> bool:
    try:
        if key not in row.keys():
            return False
        val = row[key]
        if val is None:
            return False
        return bool(int(val))
    except Exception:
        return False

def build_shifted_weekend_checker(db, employee_id: int):
    groups = db.list_groups()
    flag_map = {g["id"]: _row_flag(g, "shifted_weekends") for g in groups}
    transfers = db.conn.execute(
        "SELECT transfer_date, group_id FROM employee_transfer WHERE employee_id=? ORDER BY transfer_date DESC",
        (employee_id,),
    ).fetchall()
    history = [(d_parse(t["transfer_date"]), t["group_id"]) for t in transfers]
    emp = db.get_employee(employee_id)
    current_gid = emp["group_id"]

    def check(target_date: Optional[date] = None) -> bool:
        current = flag_map.get(current_gid, False)
        if target_date is None or not history:
            return current
        latest_date, latest_gid = history[0]
        # Простое перетаскивание не пишет приказ, но меняет текущую группу.
        if current_gid != latest_gid and target_date >= latest_date:
            return current
        for t_date, gid in history:
            if target_date >= t_date:
                return flag_map.get(gid, False)
        return current

    return check

def is_employee_shifted_weekends(db, employee_id: int, target_date: Optional[date] = None) -> bool:
    return build_shifted_weekend_checker(db, employee_id)(target_date)

def compute_month_norm_minutes(db, employee_id: int, year: int, month: int, shift_checker=None) -> int:
    if shift_checker is None: shift_checker = build_shift_checker(db, employee_id)
    shifted_checker = build_shifted_weekend_checker(db, employee_id)
    start_dt, end_dt = month_bounds_dt(year, month)
    cur, last = start_dt.date(), (end_dt - timedelta(days=1)).date()
    work_map = db.get_calendar_month(d_iso(cur), d_iso(last))

    # Собираем праздники текущего месяца И первого дня следующего месяца
    # Это нужно чтобы поймать случай: праздник 1-го числа следующего месяца,
    # а предпраздничный день — последний день текущего месяца
    next_day_after_month = end_dt.date()  # end_dt это уже первый день следующего месяца
    holidays_extended = db.get_holidays_month(d_iso(cur), d_iso(next_day_after_month))

    # Теперь вычисляем предпраздничные дни вручную с учётом межмесячной границы
    pre_holidays_set = set()
    for h_date in holidays_extended:
        candidate = h_date - timedelta(days=1)

        # Кандидат должен попадать именно в ТЕКУЩИЙ месяц
        if not (cur <= candidate <= last):
            continue

        # Проверяем что кандидат — рабочий не-праздничный день
        row = db.conn.execute(
            "SELECT is_working, is_holiday FROM calendar_day WHERE date=?",
            (d_iso(candidate),)
        ).fetchone()

        if row:
            is_working = resolve_is_working(
                candidate,
                shifted_checker(candidate),
                {candidate: bool(int(row["is_working"]))},
                holidays_extended,
            )
            is_holiday = bool(int(row["is_holiday"] if row["is_holiday"] is not None else 0))
        else:
            is_working = default_is_working(candidate, shifted_checker(candidate))
            is_holiday = False

        if is_working and not is_holiday:
            pre_holidays_set.add(candidate)

    # Собираем праздники только текущего месяца (для исключения из нормы)
    holidays_set = db.get_holidays_month(d_iso(cur), d_iso(last))

    # Собираем даты компенсаций (отгулов)
    comp_dates = set()
    rows = db.conn.execute(
        "SELECT event_date FROM compensation WHERE employee_id=? AND method='day_off' "
        "AND event_date IS NOT NULL AND event_date >= ? AND event_date <= ?",
        (employee_id, d_iso(cur), d_iso(last))
    ).fetchall()
    for r in rows:
        comp_dates.add(d_parse(r["event_date"]))

    rows_d = db.conn.execute(
        "SELECT day_off_date FROM comp_day_off_date WHERE employee_id=? "
        "AND day_off_date >= ? AND day_off_date <= ?",
        (employee_id, d_iso(cur), d_iso(last))
    ).fetchall()
    for r in rows_d:
        comp_dates.add(d_parse(r["day_off_date"]))

    total_minutes = 0

    while cur <= last:
        is_working = resolve_is_working(cur, shifted_checker(cur), work_map, holidays_set)
        is_holiday = cur in holidays_set

        # Считаем только рабочие не-праздничные дни для сменщика
        if is_working and not is_holiday and shift_checker(cur):
            status_row = db.conn.execute(
                "SELECT status FROM employee_day_status WHERE employee_id=? AND date=?",
                (employee_id, d_iso(cur))
            ).fetchone()
            has_status = status_row is not None
            has_comp = cur in comp_dates

            if cur in pre_holidays_set:
                if not has_status and not has_comp:
                    # Предпраздничный рабочий день — 7 часов
                    total_minutes += 7 * 60
                # Если на больничном/отпуске/командировке/отгуле — не добавляем ничего
            else:
                if not has_status and not has_comp:
                    # Обычный рабочий день — 8 часов
                    total_minutes += 8 * 60

        cur += timedelta(days=1)

    return total_minutes

def _get_accruals_for_period(db, employee_id, start_dt, end_dt, shift_checker, holidays_set):
    duties = db.list_duties_for_period(employee_id, start_dt, end_dt)
    breaks_map = db.breaks_for_duty_ids([int(d["id"]) for d in duties])
    shifted_checker = build_shifted_weekend_checker(db, employee_id)
    res = {"night": 0, "overtime_acc": 0, "days": 0, "shift_night": 0, "shift_holiday": 0}
    counted_days = set()

    curr_m = datetime(start_dt.year, start_dt.month, 1)
    while curr_m < end_dt:
        m_s, m_e = month_bounds_dt(curr_m.year, curr_m.month)
        c_s, c_e = max(m_s, start_dt), min(m_e, end_dt)
        norm = compute_month_norm_minutes(db, employee_id, curr_m.year, curr_m.month, shift_checker)
        m_shift = 0 
        m_duties = [d for d in duties if dt_parse(d["start_dt"]) < c_e and dt_parse(d["end_dt"]) > c_s]
        
        for d in m_duties:
            ds, de = max(dt_parse(d["start_dt"]), c_s), min(dt_parse(d["end_dt"]), c_e)
            is_emp_shifter = shift_checker(ds.date())
            is_duty_marked_as_shift = bool(int(d["is_shift"] or 0))

            for ps, pe in subtract_intervals((ds, de), breaks_map.get(int(d["id"]), [])):
                
                if is_emp_shifter:
                    if is_duty_marked_as_shift:
                        # Смена ПО ГРАФИКУ: идет в сравнение с нормой
                        m_shift += round((pe - ps).total_seconds() / 60)
                        
                        # Ночные внутри смены (используем round!)
                        for ns, ne in extract_night_intervals(ps, pe):
                            res["shift_night"] += round((ne - ns).total_seconds() / 60)
                            
                        # Праздничные внутри смены
                        cd = ps.date()
                        # Используем строгий < для pe, чтобы не захватывать лишний день в полночь
                        while cd <= pe.date():
                            d0 = datetime.combine(cd, time.min)
                            d1 = d0 + timedelta(days=1)
                            if cd in holidays_set:
                                inter = intersect(ps, pe, d0, d1)
                                if inter: 
                                    res["shift_holiday"] += round((inter[1] - inter[0]).total_seconds() / 60)
                            cd += timedelta(days=1)
                    else:
                        # ВНЕ ГРАФИКА: Сверхурочка не копится, даем Дни
                        for ns, ne in extract_night_intervals(ps, pe):
                            res["night"] += round((ne - ns).total_seconds() / 60)
                            
                        cd = ps.date()
                        while cd <= pe.date():
                            if intersect(ps, pe, datetime.combine(cd, time(6,0)), datetime.combine(cd, time(22,0))):
                                res["days"] += 1
                            cd += timedelta(days=1)

                else:
                    # ПЯТИДНЕВКА
                    for ns, ne in extract_night_intervals(ps, pe):
                        res["night"] += round((ne - ns).total_seconds() / 60)

                    cd = ps.date()
                    while cd <= pe.date():
                        is_w = db.get_calendar_month(d_iso(cd), d_iso(cd)).get(cd, default_is_working(cd, shifted_checker(cd)))
                        if not is_w or cd in holidays_set:
                            if intersect(ps, pe, datetime.combine(cd, time(6,0)), datetime.combine(cd, time(22,0))):
                                counted_days.add(cd)
                        cd += timedelta(days=1)
                        
        if norm > 0 or m_shift > 0: 
            res["overtime_acc"] += (m_shift - norm)
            
        curr_m = m_e

    res["days"] += len(counted_days)
    return res

def compute_month_summary(db, employee_id: int, year: int, month: int) -> dict:
    shift_checker = build_shift_checker(db, employee_id)
    emp = db.get_employee(employee_id)
    hire_y, hire_m = safe_get_hire_date(emp["start_month"])
    
    if year < hire_y or (year == hire_y and month < hire_m):
        return {k: 0 for k in ["norm_minutes", "shift_minutes", "shift_night_minutes", "shift_holiday_minutes", "start_hours", "start_overtime", "start_days", "acc_hours", "acc_overtime", "acc_days", "comp_hours", "comp_overtime", "comp_days", "end_hours", "end_overtime", "end_days"]} | {"is_shift": False}

    m_start, m_end = month_bounds_dt(year, month)
    m_s_iso, m_e_iso = d_iso(m_start.date()), d_iso(m_end.date())
    y_start = datetime(year, hire_m, 1) if year == hire_y else datetime(year, 1, 1)

    # 1. ТЕКУЩИЙ ГОД (База накоплений)
    if year == hire_y:
        base_h, base_o, base_d = int(emp["opening_minutes"] or 0), int(emp["opening_overtime_minutes"] or 0), int(emp["opening_days"] or 0)
    else:
        # Для последующих лет баланс переходит из декабря прошлого года
        # Рекурсия здесь безопасна, так как глубина - всего несколько лет
        prev_summ = compute_month_summary(db, employee_id, year - 1, 12)
        base_h, base_o, base_d = prev_summ["end_hours"], prev_summ["end_overtime"], prev_summ["end_days"]

    # 2. ЭТАЛОН ПРОШЛОГО ГОДА (Заначка)
    ph, po, pd = int(emp["prev_opening_minutes"] or 0), int(emp["prev_opening_overtime_minutes"] or 0), int(emp["prev_opening_days"] or 0)

    # Вспомогательная функция для списаний из Эталона (1900 год)
    def get_prev_year_spent(s_iso, e_iso):
        # Часы и сверхнорма — ищем по order_date среди записей с меткой 1900
        r_h = db.conn.execute("""
            SELECT unit, SUM(amount_minutes) as sm 
            FROM compensation 
            WHERE employee_id=? 
              AND event_date='1900-01-01' 
              AND unit IN ('hours', 'overtime')
              AND order_date >= ? AND order_date < ? 
            GROUP BY unit
        """, (employee_id, s_iso, e_iso)).fetchall()

        # Дни через отгулы (comp_day_off_date) — для day_off метода
        r_d_dayoff = db.conn.execute("""
            SELECT COUNT(*) as cd 
            FROM comp_day_off_date 
            WHERE employee_id=? 
              AND day_off_date >= ? AND day_off_date < ?
              AND compensation_id IN (
                  SELECT id FROM compensation 
                  WHERE event_date='1900-01-01' 
                    AND method='day_off'
                    AND unit='days'
              )
        """, (employee_id, s_iso, e_iso)).fetchone()["cd"] or 0

        # Дни через приказ (amount_days) — для money метода
        # Здесь смотрим по order_date, так как у денежных нет comp_day_off_date
        r_d_money = db.conn.execute("""
            SELECT COALESCE(SUM(amount_days), 0) as sd
            FROM compensation
            WHERE employee_id=?
              AND event_date='1900-01-01'
              AND method='money'
              AND unit='days'
              AND order_date >= ? AND order_date < ?
        """, (employee_id, s_iso, e_iso)).fetchone()["sd"] or 0

        res = {
            "hours": 0, 
            "overtime": 0, 
            "days": int(r_d_dayoff) + int(r_d_money)
        }
        for row in r_h:
            if row["unit"] == "hours": res["hours"] = int(row["sm"] or 0)
            elif row["unit"] == "overtime": res["overtime"] = int(row["sm"] or 0)
        return res

    spent_before = get_prev_year_spent("1900-01-01", m_s_iso)
    spent_now = get_prev_year_spent(m_s_iso, m_e_iso)

    # 3. СПИСАНИЯ ТЕКУЩЕГО ГОДА (Реальные)
    def get_comp_real(s_iso, e_iso):
        r_m = db.conn.execute("""
            SELECT unit, SUM(amount_minutes) as sm 
            FROM compensation 
            WHERE employee_id=? 
            AND (event_date IS NOT NULL AND event_date != '' AND event_date != '1900-01-01')
            AND event_date >= ? 
            AND event_date < ? 
            GROUP BY unit
        """, (employee_id, s_iso, e_iso)).fetchall()

        r_d1 = db.conn.execute("""
            SELECT SUM(amount_days) as sd 
            FROM compensation 
            WHERE employee_id=? 
            AND unit='days' AND method='money' 
            AND (event_date IS NOT NULL AND event_date != '' AND event_date != '1900-01-01')
            AND event_date >= ? 
            AND event_date < ?
        """, (employee_id, s_iso, e_iso)).fetchone()["sd"] or 0

        r_d2 = db.conn.execute("""
            SELECT COUNT(*) as cd 
            FROM comp_day_off_date 
            WHERE employee_id=? 
            AND day_off_date >= ? AND day_off_date < ? 
            AND compensation_id IN (
                SELECT id FROM compensation 
                WHERE event_date IS NOT NULL 
                AND event_date != ''
                AND event_date != '1900-01-01'
            )
        """, (employee_id, s_iso, e_iso)).fetchone()["cd"] or 0

        res = {"hours": 0, "overtime": 0, "days": int(r_d1 or 0) + int(r_d2 or 0)}
        for row in r_m:
            if row["unit"] == "hours": res["hours"] = int(row["sm"] or 0)
            elif row["unit"] == "overtime": res["overtime"] = int(row["sm"] or 0)
        return res

    comp_before = get_comp_real(d_iso(y_start.date()), m_s_iso)
    comp_now = get_comp_real(m_s_iso, m_e_iso)

    # Накопления
    ytd_before = _get_accruals_for_period(db, employee_id, y_start, m_start, shift_checker, db.get_holidays_month(d_iso(date(year, 1, 1)), d_iso(date(year, 12, 31))))
    this_acc = _get_accruals_for_period(db, employee_id, m_start, m_end, shift_checker, db.get_holidays_month(d_iso(date(year, 1, 1)), d_iso(date(year, 12, 31))))
    norm_m = compute_month_norm_minutes(db, employee_id, year, month, shift_checker)

    # 4. РАСЧЕТ ИТОГОВ (МАТЕМАТИКА)
    
    # Остатки Текущего Года (ЧЕРНЫЕ)
    start_h = base_h + ytd_before["night"] - comp_before["hours"]
    start_o = base_o + ytd_before["overtime_acc"] - comp_before["overtime"]
    start_d = base_d + ytd_before["days"] - comp_before["days"]
    
    # !!! ИСПРАВЛЕНИЕ: Вычитаем только comp_now (НЕ spent_now) !!!
    end_h = start_h + this_acc["night"] - comp_now["hours"]
    end_o = start_o + this_acc["overtime_acc"] - comp_now["overtime"]
    end_d = start_d + this_acc["days"] - comp_now["days"]

    # Остатки Прошлого Года (ЗЕЛЕНЫЕ)
    prev_h_start = ph - spent_before["hours"]
    prev_o_start = po - spent_before["overtime"]
    prev_d_start = pd - spent_before["days"]
    
    prev_h_end = prev_h_start - spent_now["hours"]
    prev_o_end = prev_o_start - spent_now["overtime"]
    prev_d_end = prev_d_start - spent_now["days"]

    return {
        "norm_minutes": norm_m,
        "shift_minutes": this_acc["night"] if not shift_checker() else (norm_m + (this_acc["overtime_acc"] if norm_m > 0 else 0)),
        "is_shift": shift_checker(),
        "start_hours": start_h, "start_overtime": start_o, "start_days": start_d,
        "acc_hours": this_acc["night"], "acc_overtime": this_acc["overtime_acc"], "acc_days": this_acc["days"],
        "shift_night": this_acc["shift_night"], "shift_holiday": this_acc["shift_holiday"],
        
        # РАЗДЕЛЬНЫЕ СПИСАНИЯ
        "comp_h_real": comp_now["hours"], # Только этот год
        "comp_h_prev": spent_now["hours"], # Только прошлый
        "comp_o_real": comp_now["overtime"],
        "comp_o_prev": spent_now["overtime"],
        "comp_d_real": comp_now["days"],
        "comp_d_prev": spent_now["days"],
        "comp_hours": comp_now["hours"] + spent_now["hours"],
        "comp_overtime": comp_now["overtime"] + spent_now["overtime"],
        "comp_days": comp_now["days"] + spent_now["days"],

        "end_hours": end_h, "end_overtime": end_o, "end_days": end_d,
        "prev_h_start": prev_h_start, "prev_o_start": prev_o_start, "prev_d_start": prev_d_start,
        "prev_h_end": prev_h_end, "prev_o_end": prev_o_end, "prev_d_end": prev_d_end
    }

def validate_non_negative_over_year(db, employee_id: int, year: int) -> tuple[bool, str]:
    emp = db.get_employee(employee_id)
    hy, hm = safe_get_hire_date(emp["start_month"])
    shift_checker = build_shift_checker(db, employee_id)
    
    months_names = ["январе", "феврале", "марте", "апреле", "мае", "июне", 
                    "июле", "августе", "сентябре", "октябре", "ноябре", "декабре"]
    
    for m in range(1, 13):
        if year < hy or (year == hy and m < hm):
            continue
            
        s = compute_month_summary(db, employee_id, year, m)
        
        # Для сменщиков переработка может быть отрицательной внутри года —
        # это нормально, смены распределяются неравномерно.
        # Проверяем переработку только для пятидневщиков.
        is_shift_month = shift_checker(date(year, m, 1))
        
        if s["end_hours"] < 0: 
            return False, f"Ошибка: Не хватает ночных часов в {months_names[m-1]}."
        if not is_shift_month and s["end_overtime"] < 0: 
            return False, f"Ошибка: Не хватает переработки в {months_names[m-1]}."
        if s["end_days"] < 0: 
            return False, f"Ошибка: Не хватает дней отгула в {months_names[m-1]}."

        if s["prev_h_end"] < 0: 
            return False, f"Ошибка: В {months_names[m-1]} превышен лимит остатков прошлого года (Ночные)."
        if s["prev_o_end"] < 0: 
            return False, f"Ошибка: В {months_names[m-1]} превышен лимит остатков прошлого года (Сверх нормы)."
        if s["prev_d_end"] < 0: 
            return False, f"Ошибка: В {months_names[m-1]} превышен лимит остатков прошлого года (Дни)."
            
    return True, ""