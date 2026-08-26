import sqlite3
import uuid
from datetime import date, datetime, timedelta
from typing import Optional, Any
from contextlib import contextmanager
from utils import * # Подключаем наши формулы!

class DBError(Exception): pass
class CalendarMissingError(DBError): pass

class DB:
    def __init__(self, path: str):
        self.path = path
        
        # МАГИЯ: Увеличиваем таймаут, чтобы SQLite не блокировал Питон
        self.conn = sqlite3.connect(path, timeout=10.0) 
        self.conn.row_factory = sqlite3.Row
        
        # --- ТУРБО-РЕЖИМ (УБИРАЕТ ЛАГИ ЧТЕНИЯ/ЗАПИСИ) ---
        self.conn.execute("PRAGMA journal_mode = WAL;") 
        self.conn.execute("PRAGMA synchronous = NORMAL;") 
        self.conn.execute("PRAGMA temp_store = MEMORY;") 
        # ------------------------------------------------
        
        self.conn.execute("PRAGMA foreign_keys = ON;")
        
        # ВСТАВЛЯЙ ЭТО СЮДА:
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN prev_opening_minutes INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN prev_opening_overtime_minutes INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN prev_opening_days INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        # ---------------------
        
        self._undo_stack = []
        self._redo_stack = []
        self._max_history = 20
        
        self._init_or_migrate()

        # Подорожники для старых баз
        try: self.conn.execute("ALTER TABLE employee_group ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee_group ADD COLUMN is_shift INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee_group ADD COLUMN shifted_weekends INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE duty ADD COLUMN is_shift INTEGER NOT NULL DEFAULT 0")
        except Exception: pass   
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN opening_overtime_minutes INTEGER NOT NULL DEFAULT 0")
        except Exception: pass        
        try: self.conn.execute("ALTER TABLE calendar_day ADD COLUMN is_holiday INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN prev_opening_minutes INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN prev_opening_overtime_minutes INTEGER NOT NULL DEFAULT 0")
        except Exception: pass
        try: self.conn.execute("ALTER TABLE employee ADD COLUMN prev_opening_days INTEGER NOT NULL DEFAULT 0")
        except Exception: pass

        

        # --- СОЗДАНИЕ ТАБЛИЦЫ ДЛЯ ИСТОРИИ ПЕРЕВОДОВ ---
        try: 
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS employee_transfer (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, 
                    employee_id INTEGER, 
                    transfer_date TEXT, 
                    group_id INTEGER
                )
            """)
        except Exception as e: 
            print(f"Ошибка создания таблицы переводов: {e}")

        # --- ФИКС СТАРЫХ БАЗ: компенсации без event_date ---
        self._migrate_fix_null_event_dates()

    def save_snapshot(self) -> None:
        # Делаем мгновенную бинарную копию базы в оперативную память для Ctrl+Z
        mem_db = sqlite3.connect(":memory:")
        mem_db.row_factory = sqlite3.Row
        self.conn.backup(mem_db)
        
        self._undo_stack.append(mem_db)
        if len(self._undo_stack) > self._max_history:
            old = self._undo_stack.pop(0)
            old.close()
            
        for db in self._redo_stack:
            db.close()
        self._redo_stack.clear()

    def begin(self) -> None:
        """Умный старт транзакции: сначала делаем снимок, потом открываем запись"""
        self.save_snapshot()
        self.conn.execute("BEGIN;")

    @contextmanager
    def transaction(self):
        """
        Умный менеджер. Сам делает COMMIT при успехе
        и ROLLBACK при любой ошибке.
        """
        self.begin() # Запускаем транзакцию и делаем снимок для Ctrl+Z
        try:
            yield # Здесь выполняется код из main.py
            self.conn.execute("COMMIT;")
        except Exception as e:
            self.conn.execute("ROLLBACK;")
            raise e # Пробрасываем ошибку дальше, чтобы показать красное уведомление

    def undo(self) -> bool:
        """Шаг назад (Ctrl+Z)"""
        if not self._undo_stack: 
            return False
        mem_db = sqlite3.connect(":memory:")
        self.conn.backup(mem_db)
        self._redo_stack.append(mem_db)
        prev_db = self._undo_stack.pop()
        prev_db.backup(self.conn)
        prev_db.close()
        return True

    def redo(self) -> bool:
        """Шаг вперед (Ctrl+Y)"""
        if not self._redo_stack: 
            return False
        mem_db = sqlite3.connect(":memory:")
        self.conn.backup(mem_db)
        self._undo_stack.append(mem_db)
        next_db = self._redo_stack.pop()
        next_db.backup(self.conn)
        next_db.close()
        return True

    def close(self) -> None:
        self.conn.close()

    def _migrate_fix_null_event_dates(self) -> None:
        """
        Фикс для старых баз: компенсации типа day_off могли записываться
        без event_date. Новая логика их не видит из-за NULL != '1900-01-01'.
        """
        # Для отгулов с датами в comp_day_off_date — берём минимальную дату отгула
        self.conn.execute("""
            UPDATE compensation
            SET event_date = (
                SELECT MIN(day_off_date) 
                FROM comp_day_off_date 
                WHERE compensation_id = compensation.id
            )
            WHERE method = 'day_off'
            AND (event_date IS NULL OR event_date = '')
            AND id IN (SELECT DISTINCT compensation_id FROM comp_day_off_date)
        """)

        # Для одиночных hours-отгулов совсем без дат — ставим заглушку
        self.conn.execute("""
            UPDATE compensation
            SET event_date = '1970-01-01'
            WHERE method = 'day_off'
            AND (event_date IS NULL OR event_date = '')
            AND id NOT IN (SELECT DISTINCT compensation_id FROM comp_day_off_date)
        """)

        self.conn.commit()

    def _has_table(self, name: str) -> bool:
        r = self.conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone()
        return r is not None

    def _init_or_migrate(self) -> None:
        if not self._has_table("meta"):
            self._init_schema()
            return
        # Миграции пока пропустим для экономии места, они тут есть в твоем оригинальном коде

    def _init_schema(self) -> None:
        c = self.conn
        c.execute("BEGIN;")
        try:
            c.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);")
            c.execute(
                """
                CREATE TABLE department_settings (
                    id INTEGER PRIMARY KEY CHECK (id=1),
                    department_name TEXT NOT NULL,
                    resp_position TEXT,
                    resp_rank TEXT,
                    resp_last_name TEXT,
                    resp_first_name TEXT,
                    resp_middle_name TEXT
                );
                """
            )
            c.execute(
                """
                CREATE TABLE employee_group (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    is_shift INTEGER NOT NULL DEFAULT 0,
                    shifted_weekends INTEGER NOT NULL DEFAULT 0
                );
                """
            )
            c.execute(
                """
                CREATE TABLE employee (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    group_id INTEGER REFERENCES employee_group(id) ON DELETE SET NULL,
                    last_name TEXT NOT NULL,
                    first_name TEXT NOT NULL,
                    middle_name TEXT,
                    rank TEXT,
                    position TEXT,
                    start_month TEXT NOT NULL,
                    end_date TEXT,
                    end_reason TEXT,
                    opening_minutes INTEGER NOT NULL DEFAULT 0,
                    opening_overtime_minutes INTEGER NOT NULL DEFAULT 0,
                    opening_days INTEGER NOT NULL DEFAULT 0,
                    prev_opening_minutes INTEGER NOT NULL DEFAULT 0,
                    prev_opening_overtime_minutes INTEGER NOT NULL DEFAULT 0,
                    prev_opening_days INTEGER NOT NULL DEFAULT 0,
                    sort_order INTEGER NOT NULL DEFAULT 0
                );
                """
            )
            c.execute(
                """
                CREATE TABLE duty (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_id INTEGER NOT NULL REFERENCES employee(id) ON DELETE CASCADE,
                    start_dt TEXT NOT NULL,
                    end_dt TEXT NOT NULL,
                    comment TEXT,
                    is_shift INTEGER NOT NULL DEFAULT 0
                );
                """
            )
            c.execute(
                """
                CREATE TABLE duty_break (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    duty_id INTEGER NOT NULL REFERENCES duty(id) ON DELETE CASCADE,
                    start_dt TEXT NOT NULL,
                    end_dt TEXT NOT NULL
                );
                """
            )            
            c.execute(
                """
                CREATE TABLE compensation (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_id INTEGER NOT NULL REFERENCES employee(id) ON DELETE CASCADE,
                    unit TEXT NOT NULL,
                    method TEXT NOT NULL,
                    event_date TEXT,
                    amount_minutes INTEGER,
                    amount_days INTEGER,
                    order_no TEXT,
                    order_date TEXT,
                    comment TEXT
                );
                """
            )
            c.execute(
                """
                CREATE TABLE comp_day_off_date (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    compensation_id INTEGER NOT NULL REFERENCES compensation(id) ON DELETE CASCADE,
                    employee_id INTEGER NOT NULL REFERENCES employee(id) ON DELETE CASCADE,
                    day_off_date TEXT NOT NULL,
                    UNIQUE(employee_id, day_off_date)
                );
                """
            )
            c.execute(
                """
                CREATE TABLE calendar_day (
                    date TEXT PRIMARY KEY,
                    is_working INTEGER NOT NULL,
                    is_holiday INTEGER NOT NULL DEFAULT 0
                );
                """
            )
            c.execute(
                """
                CREATE TABLE employee_day_status (
                    employee_id INTEGER NOT NULL REFERENCES employee(id) ON DELETE CASCADE,
                    date TEXT NOT NULL,
                    status TEXT NOT NULL,
                    UNIQUE(employee_id, date)
                );
                """
            )
            c.execute(
                """
                CREATE TABLE employee_transfer (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, 
                    employee_id INTEGER NOT NULL REFERENCES employee(id) ON DELETE CASCADE, 
                    transfer_date TEXT NOT NULL, 
                    group_id INTEGER REFERENCES employee_group(id) ON DELETE SET NULL
                );
                """
            )

            # --- ДОБАВЛЯЕМ ИНДЕКСЫ ДЛЯ СКОРОСТИ ---
            c.execute("CREATE INDEX IF NOT EXISTS idx_duty_dates ON duty (employee_id, start_dt, end_dt);")
            c.execute("CREATE INDEX IF NOT EXISTS idx_comp_dates ON compensation (employee_id, event_date);")
            c.execute("CREATE INDEX IF NOT EXISTS idx_emp_group ON employee (group_id);")
            c.execute("CREATE INDEX IF NOT EXISTS idx_emp_transfer ON employee_transfer (employee_id, transfer_date);")
            # --------------------------------------

            # SCHEMA_VERSION = 5 (или какая у тебя задана в utils.py)
            self.set_meta("schema_version", "5")
            self.set_meta("db_uuid", str(uuid.uuid4()))
            self.set_meta("created_at", datetime.now().isoformat(timespec="seconds"))
            c.execute("INSERT INTO department_settings (id, department_name) VALUES (1, ?)", ("Подразделение",))
            c.execute("COMMIT;")
        except Exception as e:
            c.execute("ROLLBACK;")
            print(f"ОШИБКА СОЗДАНИЯ СХЕМЫ БАЗЫ ДАННЫХ: {e}")
            raise

    def get_meta(self, key: str, default: Optional[str] = None) -> Optional[str]:
        r = self.conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return r["value"] if r else default

    def set_meta(self, key: str, value: str) -> None:
        self.conn.execute("INSERT INTO meta(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))

    def get_department_name(self) -> str:
        r = self.conn.execute("SELECT department_name FROM department_settings WHERE id=1").fetchone()
        return r["department_name"] if r else "Подразделение"

    def get_department_settings(self) -> sqlite3.Row:
        return self.conn.execute("SELECT * FROM department_settings WHERE id=1").fetchone()

    def update_department_settings(self, **fields: Any) -> None:
        if not fields:
            return
        cols = ", ".join([f"{k}=?" for k in fields.keys()])
        vals = list(fields.values())
        self.conn.execute(f"UPDATE department_settings SET {cols} WHERE id=1", vals)

    # --- ГРУППЫ (То, что нам нужно прямо сейчас) ---
    def list_groups(self) -> list[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM employee_group ORDER BY sort_order, id").fetchall()

    def add_group(self, name: str, is_shift: bool = False, shifted_weekends: bool = False) -> int:
        cur = self.conn.execute(
            "INSERT INTO employee_group(name, is_shift, shifted_weekends) VALUES (?, ?, ?)",
            (name.strip(), int(is_shift), int(shifted_weekends)),
        )
        return int(cur.lastrowid)

    def delete_group(self, group_id: int) -> None:
        self.conn.execute("DELETE FROM employee_group WHERE id=?", (group_id,))

    def set_group_shifted_weekends(self, group_id: int, enabled: bool) -> None:
        self.conn.execute(
            "UPDATE employee_group SET shifted_weekends=? WHERE id=?",
            (int(bool(enabled)), int(group_id)),
        )

    def set_employee_group(self, employee_id: int, group_id: Optional[int]) -> None:
        self.conn.execute("UPDATE employee SET group_id=? WHERE id=?", (group_id, employee_id))    
        
    def list_employees_for_month(self, year: int, month: int, active_only: bool, search: str = "") -> list[sqlite3.Row]:
        m = f"{year:04d}-{month:02d}"
        params: list[Any] = []
        where = []
        if active_only:
            where.append("e.start_month <= ?")
            where.append("(e.end_date IS NULL OR substr(e.end_date,1,7) >= ?)")
            params += [m, m]
            
        sql = """
            SELECT e.* 
            FROM employee e
            LEFT JOIN employee_group eg ON e.group_id = eg.id
        """
        if where:
            sql += " WHERE " + " AND ".join(where)
            
        # Сначала порядок группы, потом порядок сотрудника внутри группы, потом алфавит
        sql += " ORDER BY COALESCE(eg.sort_order, 9999), e.sort_order, e.last_name, e.first_name"
        return self.conn.execute(sql, params).fetchall()

    def get_employee(self, employee_id: int) -> sqlite3.Row:
        r = self.conn.execute("SELECT * FROM employee WHERE id=?", (employee_id,)).fetchone()
        if not r:
            raise DBError("Сотрудник не найден.")
        return r

    # --- НОВЫЕ МЕТОДЫ ДЛЯ КАЛЕНДАРЯ ---
    def get_calendar_month(self, start_date_iso: str, end_date_iso: str) -> dict[date, bool]:
        """Возвращает словарь {дата: рабочий_ли_день}"""
        rows = self.conn.execute("SELECT date, is_working FROM calendar_day WHERE date >= ? AND date <= ?", (start_date_iso, end_date_iso)).fetchall()
        return {d_parse(r["date"]): bool(int(r["is_working"])) for r in rows}

    def get_statuses_for_period(self, employee_id: int, start_date_iso: str, end_date_iso: str) -> dict[date, str]:
        rows = self.conn.execute("SELECT date, status FROM employee_day_status WHERE employee_id=? AND date>=? AND date<=?", (employee_id, start_date_iso, end_date_iso)).fetchall()
        return {d_parse(r["date"]): r["status"] for r in rows}

    def list_duties_for_period(self, employee_id: int, start_dt: datetime, end_dt: datetime) -> list[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM duty WHERE employee_id=? AND end_dt > ? AND start_dt < ? ORDER BY start_dt", (employee_id, dt_iso(start_dt), dt_iso(end_dt))).fetchall()

    def breaks_for_duty_ids(self, duty_ids: list[int]) -> dict[int, list[tuple[datetime, datetime]]]:
        out = {}
        if not duty_ids: return out
        ph = ",".join(["?"] * len(duty_ids))
        rows = self.conn.execute(f"SELECT duty_id, start_dt, end_dt FROM duty_break WHERE duty_id IN ({ph}) ORDER BY duty_id, start_dt", duty_ids).fetchall()
        for r in rows:
            out.setdefault(int(r["duty_id"]), []).append((dt_parse(r["start_dt"]), dt_parse(r["end_dt"])))
        return out

    def list_compensations_for_period(self, employee_id: int, start_date_iso: str, end_date_iso: str) -> list[sqlite3.Row]:
        """Ищет все компенсации, которые должны быть видны в календаре за период"""
        return self.conn.execute("""
            SELECT * FROM compensation WHERE employee_id=? AND method<>'money' AND (
                (event_date >= ? AND event_date <= ?) OR 
                (order_date >= ? AND order_date <= ?) OR 
                (id IN (SELECT compensation_id FROM comp_day_off_date WHERE employee_id=? AND day_off_date >= ? AND day_off_date <= ?))
            )
        """, (employee_id, start_date_iso, end_date_iso, start_date_iso, end_date_iso, employee_id, start_date_iso, end_date_iso)).fetchall()

    def get_comp_dates(self, compensation_id: int) -> list[str]:
        rows = self.conn.execute("SELECT day_off_date FROM comp_day_off_date WHERE compensation_id=? ORDER BY day_off_date", (compensation_id,)).fetchall()
        return [r["day_off_date"] for r in rows]

    def find_overlapping_duties(self, employee_id: int, start: datetime, end: datetime, exclude_duty_id: Optional[int] = None) -> list:
        params: list[Any] = [employee_id, dt_iso(start), dt_iso(end)]
        sql = "SELECT id, start_dt, end_dt FROM duty WHERE employee_id=? AND end_dt > ? AND start_dt < ?"
        if exclude_duty_id is not None:
            sql += " AND id <> ?"
            params.append(exclude_duty_id)
        rows = self.conn.execute(sql, params).fetchall()
        return [(int(r["id"]), dt_parse(r["start_dt"]), dt_parse(r["end_dt"])) for r in rows]

    def add_duty(self, employee_id: int, start: datetime, end: datetime, comment: str, is_shift: bool = False) -> int:
        cur = self.conn.execute(
            "INSERT INTO duty(employee_id,start_dt,end_dt,comment,is_shift) VALUES (?,?,?,?,?)",
            (employee_id, dt_iso(start), dt_iso(end), comment or None, int(is_shift)),
        )
        return int(cur.lastrowid)        
        
    def delete_duty(self, duty_id: int) -> None:
        self.conn.execute("DELETE FROM duty WHERE id=?", (duty_id,))

    def replace_duty_breaks(self, duty_id: int, breaks: list[tuple[datetime, datetime]]) -> None:
        self.conn.execute("DELETE FROM duty_break WHERE duty_id=?", (int(duty_id),))
        rows = []
        for s, e in breaks or []:
            rows.append((int(duty_id), dt_iso(s), dt_iso(e)))
        if rows:
            self.conn.executemany("INSERT INTO duty_break(duty_id,start_dt,end_dt) VALUES (?,?,?)", rows)

    def add_compensation_hours_dayoff(self, employee_id: int, event_date: date, minutes: int, comment: str, unit: str = "hours") -> int:
        cur = self.conn.execute(
            "INSERT INTO compensation(employee_id,unit,method,event_date,amount_minutes,comment) VALUES (?,?,?,?,?,?)",
            (employee_id, unit, "day_off", d_iso(event_date), minutes, comment or None),
        )
        return int(cur.lastrowid)

    def add_compensation_days_dayoff(self, employee_id: int, dates: list[date], comment: str) -> int:
        cur = self.conn.execute(
            "INSERT INTO compensation(employee_id,unit,method,amount_days,comment) VALUES (?,?,?,?,?)",
            (employee_id, "days", "day_off", len(dates), comment or None),
        )
        comp_id = int(cur.lastrowid)
        for d0 in dates:
            self.conn.execute(
                "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
                (comp_id, employee_id, d_iso(d0)),
            )
        return comp_id

    def delete_compensation(self, comp_id: int) -> None:
        self.conn.execute("DELETE FROM compensation WHERE id=?", (comp_id,))        
        
    def replace_comp_dayoff_dates(self, comp_id: int, employee_id: int, dates: list[date]) -> None:
        """Перезаписывает даты для компенсации периодом, если мы удалили один день из середины"""
        # Обновляем количество дней в главной таблице
        self.conn.execute("UPDATE compensation SET amount_days=? WHERE id=?", (len(dates), comp_id))
        # Удаляем старые привязки дат
        self.conn.execute("DELETE FROM comp_day_off_date WHERE compensation_id=?", (comp_id,))
        # Записываем новые даты
        for d in dates:
            self.conn.execute(
                "INSERT INTO comp_day_off_date(compensation_id, employee_id, day_off_date) VALUES (?,?,?)",
                (comp_id, employee_id, d_iso(d))
            )

    def set_day_status(self, employee_id: int, d0: date, status: str) -> None:
        if not status:
            self.conn.execute("DELETE FROM employee_day_status WHERE employee_id=? AND date=?", (employee_id, d_iso(d0)))
        else:
            self.conn.execute(
                "INSERT INTO employee_day_status(employee_id, date, status) VALUES(?,?,?) "
                "ON CONFLICT(employee_id, date) DO UPDATE SET status=excluded.status",
                (employee_id, d_iso(d0), status)
            )
            
    def toggle_calendar_day(self, d0: date) -> None:
        # Проверяем, есть ли уже этот день в базе
        r = self.conn.execute("SELECT is_working FROM calendar_day WHERE date=?", (d_iso(d0),)).fetchone()
        
        if not r:
            # Если дня в базе еще нет, стандартно: Пн-Пт(рабочие=1), Сб-Вс(выходные=0)
            current = 1 if d0.weekday() < 5 else 0
        else:
            current = r["is_working"]
            
        # Меняем значение на противоположное
        new_val = 0 if current else 1
        
        # Сохраняем в базу
        self.conn.execute(
            "INSERT INTO calendar_day(date, is_working) VALUES(?, ?) "
            "ON CONFLICT(date) DO UPDATE SET is_working=?", 
            (d_iso(d0), new_val, new_val)
        )            

    def add_employee(self, last: str, first: str, middle: str, rank: str, position: str, start_month: str, opening_minutes: int, opening_days: int, opening_overtime: int, prev_opening_minutes: int, prev_opening_overtime: int, prev_opening_days: int, group_id: Optional[int] = None) -> int:
        cur = self.conn.execute(
            """
            INSERT INTO employee(last_name,first_name,middle_name,rank,position,start_month,
                                 opening_minutes,opening_days,opening_overtime_minutes,
                                 prev_opening_minutes, prev_opening_overtime_minutes, prev_opening_days,
                                 group_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (last, first, middle or None, rank or None, position or None, start_month, 
             opening_minutes, opening_days, opening_overtime, 
             prev_opening_minutes, prev_opening_overtime, prev_opening_days,
             group_id),
        )
        return int(cur.lastrowid)

    def delete_employee(self, employee_id: int) -> None:
        self.conn.execute("DELETE FROM employee WHERE id=?", (employee_id,))

    def update_employee(self, employee_id: int, **fields) -> None:
        if not fields:
            return
        cols = ", ".join([f"{k}=?" for k in fields.keys()])
        vals = list(fields.values()) + [employee_id]
        self.conn.execute(f"UPDATE employee SET {cols} WHERE id=?", vals)

    def add_compensation_money(self, employee_id: int, unit: str, amount_minutes: int | None, amount_days: int | None, order_no: str, order_date: date, comment: str) -> int:
        cur = self.conn.execute(
            """
            INSERT INTO compensation(employee_id,unit,method,event_date,amount_minutes,amount_days,order_no,order_date,comment)
            VALUES (?,?,?,?,?,?,?,?,?)
            """,
            (employee_id, unit, "money", d_iso(order_date), amount_minutes, amount_days, order_no, d_iso(order_date), comment or None),
        )
        return int(cur.lastrowid)        
    
    def update_group_orders(self, ordered_ids: list[int]) -> None:
        for i, gid in enumerate(ordered_ids):
            self.conn.execute("UPDATE employee_group SET sort_order=? WHERE id=?", (i, gid))

    def update_employee_orders(self, ordered_ids: list[int]) -> None:
        for i, eid in enumerate(ordered_ids):
            self.conn.execute("UPDATE employee SET sort_order=? WHERE id=?", (i, eid))    

    def get_group_at_date(self, employee_id: int, target_date: date) -> Optional[int]:
        """Возвращает ID группы, в которой был сотрудник в конкретный день"""
        # Ищем самую свежую запись ДО или В день target_date
        r = self.conn.execute(
            "SELECT group_id FROM employee_transfer WHERE employee_id=? AND transfer_date<=? ORDER BY transfer_date DESC LIMIT 1", 
            (employee_id, d_iso(target_date))
        ).fetchone()
        
        if r:
            return r["group_id"]
            
        # Если истории нет (или мы смотрим до первого перевода), возвращаем текущую группу (базовый вариант)
        emp = self.get_employee(employee_id)
        return emp["group_id"]            
    
    def get_employee_transfers(self, employee_id: int) -> list[sqlite3.Row]:
        """Получает всю историю переводов сотрудника"""
        return self.conn.execute("""
            SELECT t.id, t.transfer_date, t.group_id, g.name as group_name
            FROM employee_transfer t
            LEFT JOIN employee_group g ON t.group_id = g.id
            WHERE t.employee_id=?
            ORDER BY t.transfer_date DESC
        """, (employee_id,)).fetchall()

    def sync_employee_current_group(self, employee_id: int):
        """Синхронизирует текущую группу сотрудника с его последним переводом"""
        r = self.conn.execute(
            "SELECT group_id FROM employee_transfer WHERE employee_id=? ORDER BY transfer_date DESC LIMIT 1", 
            (employee_id,)
        ).fetchone()
        
        if r:
            self.set_employee_group(employee_id, r["group_id"])
        else:
            # Если удалили вообще всю историю, снимаем его со смен (или можно оставить как есть)
            self.set_employee_group(employee_id, None)    

    def set_calendar_day_type(self, d0: date, day_type: str) -> None:
        """day_type может быть: 'work', 'weekend', 'holiday'"""
        is_working = 1 if day_type == 'work' else 0
        is_holiday = 1 if day_type == 'holiday' else 0
        
        # excluded.is_working означает, что если день уже есть в базе, мы его обновим
        self.conn.execute(
            "INSERT INTO calendar_day(date, is_working, is_holiday) VALUES(?, ?, ?) "
            "ON CONFLICT(date) DO UPDATE SET is_working=excluded.is_working, is_holiday=excluded.is_holiday",
            (d_iso(d0), is_working, is_holiday)
        )

    def get_holidays_month(self, start_date_iso: str, end_date_iso: str) -> set[date]:
        """Возвращает набор дат, которые являются праздниками"""
        try:
            rows = self.conn.execute("SELECT date FROM calendar_day WHERE is_holiday=1 AND date >= ? AND date <= ?", (start_date_iso, end_date_iso)).fetchall()
            return {d_parse(r["date"]) for r in rows}
        except Exception:
            return set()      

    def get_pre_holidays_month(self, start_date_iso: str, end_date_iso: str) -> set[date]:
        """
        Возвращает набор предпраздничных дат в запрошенном периоде.
        Смотрит также на праздники СЛЕДУЮЩЕГО дня после периода,
        чтобы поймать случай: праздник 1-го числа, предпраздничный — последний день месяца.
        """
        try:
            from datetime import timedelta

            # Берём праздники с захватом одного дня ПОСЛЕ конца периода
            end_plus_one = d_parse(end_date_iso) + timedelta(days=1)

            rows = self.conn.execute(
                "SELECT date FROM calendar_day WHERE is_holiday=1 AND date >= ? AND date <= ?",
                (start_date_iso, d_iso(end_plus_one))
            ).fetchall()

            holidays_set = {d_parse(r["date"]) for r in rows}

            pre_holidays = set()
            for h_date in holidays_set:
                candidate = h_date - timedelta(days=1)

                # Кандидат должен быть внутри запрошенного периода
                if not (d_parse(start_date_iso) <= candidate <= d_parse(end_date_iso)):
                    continue

                # Проверяем что кандидат — рабочий не-праздничный день
                row = self.conn.execute(
                    "SELECT is_working, is_holiday FROM calendar_day WHERE date=?",
                    (d_iso(candidate),)
                ).fetchone()

                if row:
                    is_working = bool(int(row["is_working"]))
                    is_holiday = bool(int(
                        row["is_holiday"] if row["is_holiday"] is not None else 0
                    ))
                else:
                    is_working = candidate.weekday() < 5
                    is_holiday = False

                if is_working and not is_holiday:
                    pre_holidays.add(candidate)

            return pre_holidays

        except Exception as e:
            print(f"Ошибка get_pre_holidays_month: {e}")
            return set()
        