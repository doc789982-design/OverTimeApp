from datetime import date, datetime, time, timedelta
from typing import Optional, Any

SCHEMA_VERSION = 5

def dt_parse(s: str) -> datetime: return datetime.fromisoformat(s)
def dt_iso(dt: datetime) -> str: return dt.isoformat(timespec="minutes")
def d_parse(s: str) -> date: return date.fromisoformat(s)
def d_iso(d: date) -> str: return d.isoformat()

def fmt_date_iso(iso: str) -> str:
    try: return d_parse(iso).strftime("%d.%m.%Y")
    except Exception: return iso or ""

def fmt_dt_iso(iso: str) -> str:
    try: return dt_parse(iso).strftime("%d.%m.%Y %H:%M")
    except Exception: return iso or ""

def parse_month(s: str) -> tuple[int, int]:
    y, m = s.split("-")
    return int(y), int(m)

def next_month(year: int, month: int) -> tuple[int, int]:
    return (year + 1, 1) if month == 12 else (year, month + 1)

def month_bounds_dt(year: int, month: int) -> tuple[datetime, datetime]:
    start = datetime(year, month, 1, 0, 0)
    ny, nm = next_month(year, month)
    end = datetime(ny, nm, 1, 0, 0)
    return start, end

def year_bounds_dt(year: int) -> tuple[datetime, datetime]:
    return datetime(year, 1, 1, 0, 0), datetime(year + 1, 1, 1, 0, 0)

def minutes_to_hhmm(m: int) -> str:
    sign = "-" if m < 0 else ""
    m = abs(m)
    return f"{sign}{m // 60}:{m % 60:02d}"

def parse_hhmm(s: str) -> time:
    try:
        hh, mm = (s or "").strip().split(":")
        return time(int(hh), int(mm))
    except Exception: return time(0, 0)

def fmt_hhmm(t: time) -> str:
    return f"{int(t.hour):02d}:{int(t.minute):02d}"

def intersect(a0: datetime, a1: datetime, b0: datetime, b1: datetime) -> Optional[tuple[datetime, datetime]]:
    s = max(a0, b0)
    e = min(a1, b1)
    return (s, e) if s < e else None

def merge_intervals(intervals: list[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    if not intervals: return []
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for s, e in intervals[1:]:
        ps, pe = merged[-1]
        if s <= pe: merged[-1] = (ps, max(pe, e))
        else: merged.append((s, e))
    return merged

def yyyymm_from_end_date(end_date_iso: str) -> str:
    return (end_date_iso or "")[:7]
    
def subtract_intervals(
    base: tuple[datetime, datetime],
    cuts: list[tuple[datetime, datetime]],
) -> list[tuple[datetime, datetime]]:
    """base minus cuts -> список непересекающихся интервалов."""
    bs, be = base
    if bs >= be:
        return []

    if not cuts:
        return [(bs, be)]

    clipped: list[tuple[datetime, datetime]] = []
    for cs, ce in cuts:
        inter = intersect(bs, be, cs, ce)
        if inter:
            clipped.append(inter)

    if not clipped:
        return [(bs, be)]

    clipped = merge_intervals(clipped)

    out: list[tuple[datetime, datetime]] = []
    cur = bs
    for cs, ce in clipped:
        if cs > cur:
            out.append((cur, min(cs, be)))
        cur = max(cur, ce)
        if cur >= be:
            break

    if cur < be:
        out.append((cur, be))

    return [(s, e) for s, e in out if s < e]

def ru_plural(n: int, one: str, few: str, many: str) -> str:
    n = abs(int(n))
    if 11 <= (n % 100) <= 14: return many
    last = n % 10
    if last == 1: return one
    if 2 <= last <= 4: return few
    return many

def fmt_minutes_ru_words(minutes: int) -> str:
    sign = "-" if minutes < 0 else ""
    m = abs(int(minutes))
    h = m // 60
    mm = m % 60
    # Сокращенный формат: "8 ч. 30 м." или просто "8 ч."
    parts = [f"{h} ч."]
    if mm > 0: parts.append(f"{mm} м.")
    return sign + " ".join(parts)