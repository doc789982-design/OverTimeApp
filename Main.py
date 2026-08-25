import os
import tempfile
import uuid
import win32print
import sys
import json
import calendar
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from datetime import date, datetime, timedelta, time
from pathlib import Path

from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from PySide6.QtGui import QIcon, QAction
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtCore import QObject, Slot, Signal, Property, QUrl, QThread, QTimer

from database import DB
from utils import fmt_date_iso, fmt_dt_iso, d_iso, d_parse, dt_parse, dt_iso, parse_hhmm, subtract_intervals, intersect, merge_intervals, fmt_minutes_ru_words
from logic import compute_month_summary, is_employee_shift, validate_non_negative_over_year
import app_update

# ====================================================
# УМНЫЙ ПЕРЕКЛЮЧАТЕЛЬ РЕЖИМОВ (DEV / PROD)
# ====================================================
IS_FROZEN = getattr(sys, 'frozen', False)
ASSETS_DIR = Path(__file__).parent

if IS_FROZEN:
    # PROD-РЕЖИМ (Работаем как .exe)
    import resources_rc # Импортируем память только в скомпилированном виде
    QML_PATH = QUrl("qrc:/main.qml")
    EXCEL_TEMPLATE_PATH = ":/Template.xlsx"
    APP_ICON_PATH = ":/app_icon.png"
else:
    # DEV-РЕЖИМ (Работаем из редактора кода)
    QML_PATH = QUrl.fromLocalFile(str(ASSETS_DIR / "main.qml"))
    EXCEL_TEMPLATE_PATH = str(ASSETS_DIR / "Template.xlsx")
    APP_ICON_PATH = str(ASSETS_DIR / "app_icon.png")
# ====================================================

def is_emp_active_in_month(emp, year: int, month: int) -> bool:
    m = f"{year:04d}-{month:02d}"
    if emp["start_month"] > m: return False
    if emp["end_date"] is None: return True
    return (emp["end_date"] or "")[:7] >= m

# ====================================================
# ФОНОВЫЙ ПОТОК ДЛЯ ЭКСПОРТА (ЧТОБЫ НЕ ВИС ИНТЕРФЕЙС)
# ====================================================
class ExportWorker(QThread):
    # Сигнал, который поток пошлет в главное окно, когда закончит
    finished_signal = Signal(bool, str) 

    def __init__(self, db_path, year, month, template_path, out_path):
        super().__init__()
        self.db_path = db_path
        self.year = year
        self.month = month
        self.template_path = template_path
        self.out_path = out_path

    def run(self):
        try:
            from database import DB
            from export import TemplateExporter
            
            # В фоновом потоке мы открываем СВОЮ копию подключения к базе,
            # чтобы не мешать главному потоку читать данные
            temp_db = DB(self.db_path)
            
            TemplateExporter.export(
                db=temp_db, 
                year=self.year, 
                month=self.month, 
                template_path=self.template_path, 
                out_path=self.out_path
            )
            temp_db.close()
            
            # Сообщаем об успехе
            self.finished_signal.emit(True, "Экспорт в Excel успешно завершен!")
        except Exception as e:
            # Сообщаем об ошибке
            self.finished_signal.emit(False, f"Ошибка экспорта: {e}")
# ====================================================

# ====================================================
# ФОНОВЫЙ ПОТОК ДЛЯ ПЕЧАТИ EXCEL
# ====================================================
class PrintWorker(QThread):
    finished_signal = Signal(bool, str)

    def __init__(self, db_path, year, month, printer_name, copies, page_from, page_to, orientation, paper_size, collate):
        super().__init__()
        self.db_path = db_path
        self.year = year
        self.month = month
        self.printer_name = printer_name
        self.copies = copies
        self.page_from = page_from
        self.page_to = page_to
        self.orientation = orientation
        self.paper_size = paper_size
        self.collate = collate

    def run(self):
        try:
            # МАГИЯ WINDOWS: Чтобы общаться с Excel из невидимого потока, 
            # нужно инициализировать COM-интерфейс
            import pythoncom
            pythoncom.CoInitialize() 
            
            import win32com.client 
            import tempfile
            import uuid
            import os
            from pathlib import Path
            from database import DB
            from export import TemplateExporter

            # 1. Формируем временный Excel-файл
            temp_db = DB(self.db_path)
            template_path = Path(__file__).parent / "Template.xlsx"
            temp_dir = Path(tempfile.gettempdir())
            temp_excel_path = temp_dir / f"overtimetab_print_{uuid.uuid4().hex[:6]}.xlsx"

            TemplateExporter.export(
                db=temp_db, 
                year=self.year, 
                month=self.month, 
                template_path=str(template_path), 
                out_path=str(temp_excel_path)
            )
            temp_db.close()

            # 2. Открываем невидимый Excel
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False       
            excel.DisplayAlerts = False 

            try:
                wb = excel.Workbooks.Open(str(temp_excel_path))
                ws = wb.ActiveSheet
                
                # 3. Настраиваем печать
                ws.PageSetup.Orientation = 2 if self.orientation == "landscape" else 1
                ws.PageSetup.PaperSize = 8 if self.paper_size == "A3" else 9
                
                print_args = {
                    "Copies": self.copies,
                    "ActivePrinter": self.printer_name,
                    "Collate": self.collate
                }
                
                if self.page_from.strip() and self.page_to.strip():
                    try:
                        print_args["From"] = int(self.page_from)
                        print_args["To"] = int(self.page_to)
                    except ValueError:
                        pass 
                
                # 4. Отправляем на принтер!
                ws.PrintOut(**print_args)
                wb.Close(SaveChanges=False)
            finally:
                excel.Quit()
                
            if temp_excel_path.exists():
                os.remove(str(temp_excel_path))
                
            self.finished_signal.emit(True, "Документ отправлен на принтер!")
        except Exception as e:
            self.finished_signal.emit(False, f"Ошибка фоновой печати: {e}")
        finally:
            try:
                pythoncom.CoUninitialize() # Обязательно закрываем за собой COM-интерфейс
            except:
                pass
# ====================================================

class UpdateStageWorker(QThread):
    """Копирует zip/папку новой версии в pending_update, не блокируя окно."""
    finished_signal = Signal(bool, str, str)  # ok, message, version

    def __init__(self, source_path, app_dir):
        super().__init__()
        self.source_path = source_path
        self.app_dir = app_dir

    def run(self):
        try:
            ver = app_update.stage_package(Path(self.source_path), Path(self.app_dir))
            self.finished_signal.emit(True, "", ver or "")
        except Exception as e:
            self.finished_signal.emit(False, str(e), "")

# ====================================================

class Backend(QObject):
    dbListChanged = Signal()
    groupListChanged = Signal()
    employeeListChanged = Signal()
    calendarDaysChanged = Signal()
    currentPeriodChanged = Signal()
    databaseOpened = Signal()
    activeDepartmentNameChanged = Signal()
    selectedEmployeeChanged = Signal()
    monthSummaryChanged = Signal() # Сигнал для нижней панели итогов
    yearSummaryChanged = Signal()
    monthPulseChanged = Signal()
    yearlyDataChanged = Signal()
    dayDetailsChanged = Signal()
    dayCompsChanged = Signal()
    moneyCompsChanged = Signal()
    showToast = Signal(str, str)    
    timeInputModeChanged = Signal()    
    hotkeysListChanged = Signal()
    printerListChanged = Signal()
    departmentDataChanged = Signal()
    employeeTransferHistoryChanged = Signal()
    isDarkThemeChanged = Signal()    
    itemDeleted = Signal(str, str)
    startHiddenChanged = Signal()
    reminderEnabledChanged = Signal()
    updateReadyChanged = Signal()
    updateBusyChanged = Signal()
    updateVersionChanged = Signal()
    updateStatusTextChanged = Signal()
    appVersionChanged = Signal()

    def __init__(self, start_hidden=False):
        super().__init__()
        self._start_hidden = start_hidden
        self._db_list = []
        self._group_list = []
        self._employee_list = []
        self._calendar_days = []
        self._month_summary = {} # Тут будут лежать итоги
        self._clipboard = None
        self._department_data = {}
        self._hotkeys_list = []
        self._day_duties = []
        self._day_comps = []  
        self._money_comps = []
        self._yearly_data = []
        self._month_pulse = {} # {1: True, 5: True} - значит в январе и мае есть данные        
        self._year_list = []
        self._time_input_mode = "slider" # По умолчанию будет ползунок      
        self._selected_employee_ratio = 0.0
        self._printer_list = []      # <--- Добавили
        self._default_printer = ""   # <--- Добавили
        self.loadPrinters()          # <--- Добавили  
        self._year_summary = {}
        self._is_dark_theme = True
        self._reminder_enabled = True    # Напоминание "сдать табель" (28-е — 5-е число)
        self._tray_hint_shown = False    # Показывали ли подсказку про работу в фоне
        # Делаем программу портативной: папка data прямо внутри папки с программой
        self.app_dir = Path(__file__).parent
        self.config_path = self.app_dir / "data" / "config.json"
        
        # Если есть старый конфиг из прошлых версий программы, копируем его сюда
        old_config = Path.home() / ".overtimetab" / "config.json"
        if old_config.exists() and not self.config_path.exists():
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            import shutil
            shutil.copy2(old_config, self.config_path)
        self.loadHotkeys()
        self.generateYearList() # <--- Сразу вызываем функцию        
        self.active_db = None
        
        self.current_year = date.today().year
        self.current_month = date.today().month
        self.current_group_id = None
        self.search_text = ""
        self.show_active_only = True
        
        self._selected_employee_id = 0
        self._active_department_name = ""
        
        self.load_databases()

        self._update_ready = False
        self._update_busy = False
        self._update_version = ""
        self._update_status = ""
        self._update_source = ""
        self._app_version = app_update.current_app_version(self.app_dir)
        self._init_updates()

    def load_databases(self):
        loaded_dbs = []
        unique_paths = set() # Сет (множество) для защиты от дубликатов!
        
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                for p in data.get("db_paths", []):
                    path_obj = Path(p)
                    # Если путь "относительный" (сохранен с флешки), склеиваем его с текущей папкой программы
                    if not path_obj.is_absolute():
                        path_obj = (self.app_dir / path_obj).resolve()
                    else:
                        path_obj = path_obj.resolve()
                    
                    # Если такой путь мы еще не добавляли и файл реально существует
                    if str(path_obj) not in unique_paths and path_obj.exists():
                        try:
                            temp_db = DB(str(path_obj))
                            dept_name = temp_db.get_department_name()
                            temp_db.close()
                            
                            loaded_dbs.append({"name": dept_name, "path": str(path_obj)})
                            unique_paths.add(str(path_obj)) # Запоминаем, что этот путь уже есть
                        except Exception: 
                            pass
                        
                # Читаем настройку интерфейса!
                ui_cfg = data.get("ui", {})
                self._time_input_mode = ui_cfg.get("time_input_mode", "slider")
                self._is_dark_theme = ui_cfg.get("theme", "dark") == "dark"
                self._reminder_enabled = ui_cfg.get("reminder_enabled", True)
                self._tray_hint_shown = ui_cfg.get("tray_hint_shown", False)
                
            except Exception as e: 
                print(f"Ошибка чтения конфига: {e}")

        self._db_list = loaded_dbs
        self.dbListChanged.emit()

    def add_to_config(self, new_path: str):
        data = {"db_paths": [], "last_db_path": None, "ui": {}}
        
        if self.config_path.exists():
            try: 
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            except: 
                pass
                
        if "db_paths" not in data: 
            data["db_paths"] = []
            
        # Нормализуем новый путь
        normalized_new_path = str(Path(new_path).resolve())

        # МАГИЯ ПОРТАТИВНОСТИ: 
        # Если база лежит внутри папки программы, обрезаем путь, чтобы он стал относительным
        try:
            rel_path = Path(normalized_new_path).relative_to(self.app_dir)
            path_to_save = str(rel_path).replace("\\", "/") # Делаем универсальные слеши
        except ValueError:
            # Если база где-то на диске C:, а программа на D:, сохраняем полный путь
            path_to_save = normalized_new_path

        # Нормализуем старые пути из конфига и чистим дубликаты
        cleaned_paths = []
        for p in data["db_paths"]:
            try:
                # Оставляем как есть, если это относительный путь, либо нормализуем
                norm_p = str(Path(p).as_posix()) if not Path(p).is_absolute() else str(Path(p).resolve())
                if norm_p not in cleaned_paths:
                    cleaned_paths.append(norm_p)
            except Exception:
                pass
                
        # Добавляем новый путь
        if path_to_save not in cleaned_paths: 
            cleaned_paths.append(path_to_save)
            
        data["db_paths"] = cleaned_paths
        data["last_db_path"] = path_to_save
        
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def _get_db_dir(self):
        """Возвращает текущую папку для сохранения баз"""
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                custom_dir = data.get("ui", {}).get("default_db_dir")
                if custom_dir and Path(custom_dir).exists():
                    return Path(custom_dir)
            except: pass
            
        # Если папка не задана, используем стандартную внутри программы
        db_dir = self.app_dir / "data" / "databases"
        db_dir.mkdir(parents=True, exist_ok=True)
        return db_dir

    @Slot()
    def loadPrinters(self):
        try:
            # Читаем все локальные и сетевые принтеры
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = [p[2] for p in win32print.EnumPrinters(flags)]
            self._printer_list = printers
            self._default_printer = win32print.GetDefaultPrinter()
            self.printerListChanged.emit()
        except Exception as e:
            print(f"Ошибка загрузки принтеров: {e}")

    @Property(list, notify=printerListChanged)
    def printerList(self): return self._printer_list

    @Property(str, notify=printerListChanged)
    def defaultPrinter(self): return self._default_printer

    @Property(list, notify=dbListChanged)
    def dbList(self): return self._db_list

    @Property(list, notify=groupListChanged)
    def groupList(self): return self._group_list

    @Property(list, notify=employeeListChanged)
    def employeeList(self): return self._employee_list

    @Property(list, notify=calendarDaysChanged)
    def calendarDays(self): return self._calendar_days

    @Property(int, notify=selectedEmployeeChanged)
    def selectedEmployeeId(self): return self._selected_employee_id

    @Property(str, notify=selectedEmployeeChanged)
    def selectedEmployeeStartMonth(self):
        if not self.active_db or self._selected_employee_id == 0:
            return "2000-01" # Значение по умолчанию, если никто не выбран
        try:
            emp = self.active_db.get_employee(self._selected_employee_id)
            return emp["start_month"] # Вернет "2026-03"
        except:
            return "2000-01"

    @Property(dict, notify=monthSummaryChanged)
    def monthSummary(self): return self._month_summary

    @Property(str, notify=currentPeriodChanged)
    def currentPeriodText(self):
        months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
        return f"{months[self.current_month - 1]} {self.current_year}"

    @Property(list, notify=dayDetailsChanged)
    def dayDuties(self): return self._day_duties

    @Property(list, notify=dayCompsChanged)
    def dayComps(self): return self._day_comps

    @Property(list, notify=moneyCompsChanged)
    def moneyComps(self): return self._money_comps

    @Property(dict, notify=departmentDataChanged)
    def departmentData(self): return self._department_data    

    @Property(list, notify=hotkeysListChanged)
    def hotkeysList(self): return self._hotkeys_list

    @Property(list, notify=yearlyDataChanged)
    def yearlyData(self): return self._yearly_data    

    @Property("QVariant", notify=monthPulseChanged)
    def monthPulse(self): return self._month_pulse

    @Property(str, notify=activeDepartmentNameChanged)
    def activeDepartmentName(self): return self._active_department_name

    yearListChanged = Signal()    

    @Property(list, notify=yearListChanged)
    def yearList(self): return self._year_list

    @Property(dict, notify=yearSummaryChanged)
    def yearSummary(self): return self._year_summary

    @Property(str, notify=timeInputModeChanged)
    def timeInputMode(self): return self._time_input_mode

    @Property(bool, notify=selectedEmployeeChanged)
    def isSelectedEmployeeShift(self):
        if not self.active_db or self._selected_employee_id == 0:
            return False
        return is_employee_shift(self.active_db, self._selected_employee_id)

    @Property(list, notify=employeeTransferHistoryChanged)
    def employeeTransferHistory(self): 
        if not hasattr(self, '_transfer_history'):
            self._transfer_history = []
        return self._transfer_history

    @Property(bool, notify=isDarkThemeChanged)
    def isDarkTheme(self): return self._is_dark_theme

    selectedEmployeeRatioChanged = Signal()

    @Property(float, notify=selectedEmployeeRatioChanged)
    def selectedEmployeeRatio(self): 
        return self._selected_employee_ratio

    @Property(bool, notify=startHiddenChanged)
    def startHidden(self): 
        return self._start_hidden

    @Slot()
    def toggleTheme(self):
        """Меняет тему и сохраняет в config.json"""
        self._is_dark_theme = not self._is_dark_theme
        try:
            data = {"db_paths": [], "last_db_path": None, "ui": {}}
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            
            ui_cfg = data.get("ui", {})
            ui_cfg["theme"] = "dark" if self._is_dark_theme else "light"
            data["ui"] = ui_cfg
            
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            print(f"Ошибка сохранения темы: {e}")
            
        self.isDarkThemeChanged.emit()

    def _write_ui_config(self, key, value):
        """Аккуратно дописывает одну настройку в секцию ui конфига (не трогая остальное)"""
        try:
            data = {"db_paths": [], "last_db_path": None, "ui": {}}
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            ui_cfg = data.get("ui", {})
            ui_cfg[key] = value
            data["ui"] = ui_cfg
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            print(f"Ошибка сохранения настройки '{key}': {e}")

    # ── Напоминание о сдаче табеля (вкл/выкл из настроек) ──────────────
    @Property(bool, notify=reminderEnabledChanged)
    def reminderEnabled(self):
        return self._reminder_enabled

    @Slot(bool)
    def setReminderEnabled(self, enabled):
        if self._reminder_enabled == bool(enabled): return
        self._reminder_enabled = bool(enabled)
        self._write_ui_config("reminder_enabled", self._reminder_enabled)
        self.reminderEnabledChanged.emit()
        self.showToast.emit(
            "Напоминания включены" if self._reminder_enabled else "Напоминания выключены",
            "success"
        )

    # ── Подсказка "программа работает в фоне" (показываем один раз) ────
    @Slot(result=bool)
    def trayHintWasShown(self):
        return self._tray_hint_shown

    @Slot()
    def setTrayHintShown(self):
        if self._tray_hint_shown: return
        self._tray_hint_shown = True
        self._write_ui_config("tray_hint_shown", True)

    def _validate_db_file(self, path: str):
        """Проверяет, что файл — настоящая база OverTimeTab (или пустой файл под новую базу).
        Возвращает кортеж (ok, текст_ошибки). Нужна, чтобы не подсунуть программе
        случайный файл (фото, документ, чужую базу) и не сломать её."""
        try:
            p = Path(path)
        except Exception:
            return False, "некорректный путь"

        if not p.exists():
            return False, "файл не найден"
        if p.is_dir():
            return False, "выбрана папка, а не файл"

        # Совсем пустой файл — разрешаем: из него создастся новая база
        if p.stat().st_size == 0:
            return True, ""

        import sqlite3 as _sq
        try:
            uri = p.resolve().as_uri() + "?mode=ro"
            con = _sq.connect(uri, uri=True)
            tables = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
            con.close()
        except Exception:
            return False, "это не файл базы данных SQLite"

        our_tables = {"employee", "duty", "employee_group", "calendar_day", "department_settings"}
        if tables and not our_tables.intersection(tables):
            return False, "это база данных другого приложения, а не табель"
        return True, ""

    @Slot(str)
    def attachDatabase(self, file_url):
        if not file_url:
            return

        # QML отдает путь в виде 'file:///C:/...', переводим в нормальный вид
        path = QUrl(file_url).toLocalFile() if file_url.startswith("file://") else file_url

        ok, err = self._validate_db_file(path)
        if not ok:
            self.showToast.emit(f"Не удалось подключить: {err}", "error")
            return

        self.add_to_config(path)
        self.load_databases()
        self.showToast.emit("База успешно подключена", "success")

    @Slot(str)
    def openDatabase(self, path):
        # Сначала убеждаемся, что файл на месте и он действительно наша база.
        # Если нет — честно ругаемся и ОСТАЕМСЯ на старой базе, а не ломаемся посередине.
        ok, err = self._validate_db_file(path)
        if not ok:
            self.showToast.emit(f"Не удалось открыть базу: {err}", "error")
            return

        try:
            new_db = DB(path)
        except Exception as e:
            self.showToast.emit(f"Не удалось открыть базу: {e}", "error")
            return

        if self.active_db:
            try: self.active_db.close()
            except Exception: pass
        self.active_db = new_db
        self._active_department_name = self.active_db.get_department_name()
        self.activeDepartmentNameChanged.emit()        
        self.add_to_config(path)
        self._selected_employee_id = 0
        self.refresh_groups()
        self.refresh_employees() 
        self.refresh_calendar()
        self.databaseOpened.emit()

    @Slot()
    def prevMonth(self):
        if self.current_month == 1: self.current_month = 12; self.current_year -= 1
        else: self.current_month -= 1
        self.refresh_yearly_panorama()
        self.refresh_calendar()
        self.refresh_employees()
        self.refresh_pulse()

    @Slot()
    def nextMonth(self):
        if self.current_month == 12: 
            self.current_month = 1
            self.current_year += 1
            self.refresh_yearly_panorama() # <--- ДОБАВИТЬ
            self.refresh_pulse()
        else: self.current_month += 1
        self.refresh_calendar()
        self.refresh_employees()

    @Slot(int)
    def jumpToMonth(self, target_month):
        if 1 <= target_month <= 12 and target_month != self.current_month:
            self.current_month = target_month
            self.refresh_calendar()
            self.refresh_employees()

    @Slot(int)
    def setGroupFilter(self, group_id):
        self.current_group_id = None if group_id == 0 else group_id
        self.refresh_employees()

    @Slot(str)
    def setSearchText(self, text):
        self.search_text = text
        self.refresh_employees()

    @Slot(bool)
    def setActiveOnly(self, is_active):
        self.show_active_only = is_active
        self.refresh_employees()

    @Slot(int)
    def selectEmployee(self, emp_id):
        self._selected_employee_id = emp_id
        self.selectedEmployeeChanged.emit()
        self.refresh_calendar() 
        self.refresh_yearly_panorama() # <--- ДОБАВИТЬ ЭТО
        self.refresh_pulse()

    @Slot(str, str, str, bool, str, str)
    def saveDuty(self, date_str, start_time_str, end_time_str, is_shift, comment, breaks_json):
        if not self.active_db or self._selected_employee_id == 0:
            return

        try:
            d0 = d_parse(date_str)
            s_time = parse_hhmm(start_time_str)
            e_time = parse_hhmm(end_time_str)

            start_dt = datetime.combine(d0, s_time)
            end_dt = datetime.combine(d0, e_time)

            if end_dt <= start_dt:
                end_dt += timedelta(days=1)

            overlaps = self.active_db.find_overlapping_duties(self._selected_employee_id, start_dt, end_dt)
            if overlaps:
                self.showToast.emit("Ошибка: Пересечение с другим дежурством", "error")
                return

            # МАГИЯ: Открываем безопасную транзакцию
            with self.active_db.transaction():
                duty_id = self.active_db.add_duty(
                    self._selected_employee_id, start_dt, end_dt, comment, is_shift
                )
                
                if breaks_json:
                    breaks_data = json.loads(breaks_json)
                    breaks_list = []
                    for b in breaks_data:
                        bs_time = time(int(b["start_h"]), int(b["start_m"]))
                        be_time = time(int(b["end_h"]), int(b["end_m"]))
                        
                        bs_dt = datetime.combine(d0, bs_time)
                        if bs_dt < start_dt: bs_dt += timedelta(days=1)
                            
                        be_dt = datetime.combine(bs_dt.date(), be_time)
                        if be_dt <= bs_dt: be_dt += timedelta(days=1)
                            
                        breaks_list.append((bs_dt, be_dt))
                        
                    self.active_db.replace_duty_breaks(duty_id, breaks_list)
                    
            self.refresh_calendar()
            self.showToast.emit("Дежурство сохранено", "success")
            
        except Exception as e:
            self.showToast.emit(f"Ошибка сохранения: {e}", "error")

    @Slot(str, str, str, str, bool)
    def saveCompensation(self, dates_csv, comp_type, amount_str, comment, use_prev_year):
        if not self.active_db or self._selected_employee_id == 0 or not dates_csv: return
        try:
            dates_list = [d_parse(d.strip()) for d in dates_csv.split(",") if d.strip()]
            with self.active_db.transaction():
                logic_date = date(1900, 1, 1) if use_prev_year else dates_list[0]

                if comp_type in ("hours", "overtime"):
                    # Часовые компенсации — и ночные и сверх нормы
                    for d0 in dates_list:
                        self.active_db.conn.execute(
                            "INSERT INTO compensation(employee_id,unit,method,event_date,order_date,amount_minutes,comment) VALUES (?,?,?,?,?,?,?)",
                            (self._selected_employee_id, comp_type, "day_off", d_iso(logic_date), d_iso(d0), int(amount_str), comment or None)
                        )
                else:
                    # Дни
                    cur = self.active_db.conn.execute(
                        "INSERT INTO compensation(employee_id,unit,method,amount_days,comment,event_date) VALUES (?,?,?,?,?,?)",
                        (self._selected_employee_id, "days", "day_off", len(dates_list), comment or None, d_iso(logic_date)),
                    )
                    comp_id = cur.lastrowid
                    for d0 in dates_list:
                        self.active_db.conn.execute(
                            "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
                            (comp_id, self._selected_employee_id, d_iso(d0))
                        )

                is_valid, err = validate_non_negative_over_year(self.active_db, self._selected_employee_id, self.current_year)
                if not is_valid:
                    raise Exception(err)

            self.refresh_calendar()
            self.refresh_yearly_panorama()
            self.refresh_pulse()
        except Exception as e:
            self.showToast.emit(str(e), "error")

    @Slot(int, result="QVariant")
    def getAvailableBalances(self, year):
        """Возвращает остаток для диалогов с учетом уже сделанных трат Эталона"""
        if not self.active_db or self._selected_employee_id == 0:
            return {"hours": 0, "overtime": 0, "days": 0}

        # Нам нужны функции из logic.py для точного расчета
        from logic import safe_get_hire_date, compute_month_summary

        emp = self.active_db.get_employee(self._selected_employee_id)
        hire_y, hire_m = safe_get_hire_date(emp["start_month"])
        
        # 1. СЛУЧАЙ: СПИСАНИЕ ИЗ ПРОШЛОГО ГОДА (ЭТАЛОНА)
        # Если запрашиваемый год меньше текущего рабочего года
        if year < self.current_year:
            # Берем чистые цифры из карточки (наш Эталон)
            h = int(emp["prev_opening_minutes"] or 0)
            o = int(emp["prev_opening_overtime_minutes"] or 0)
            d = int(emp["prev_opening_days"] or 0)
            
            # Считаем, сколько мы УЖЕ потратили из этого Эталона в ТЕКУЩЕМ году.
            # (Ищем все записи с меткой 1900, сделанные в текущем self.current_year)
            current_y_str = str(self.current_year)
            
            # Вычитаем потраченные часы и сверхнорму
            r_h = self.active_db.conn.execute("""
                SELECT unit, SUM(amount_minutes) as sm 
                FROM compensation 
                WHERE employee_id=? AND event_date='1900-01-01' AND unit IN ('hours', 'overtime')
                  AND substr(order_date, 1, 4) = ?
                GROUP BY unit
            """, (self._selected_employee_id, current_y_str)).fetchall()
            
            # Вычитаем потраченные дни
            r_d = self.active_db.conn.execute("""
                SELECT COUNT(*) as cd FROM comp_day_off_date 
                WHERE employee_id=? AND day_off_date >= ? AND day_off_date <= ?
                  AND compensation_id IN (SELECT id FROM compensation WHERE event_date='1900-01-01' AND unit='days')
            """, (self._selected_employee_id, current_y_str + "-01-01", current_y_str + "-12-31")).fetchone()["cd"] or 0
            
            for row in r_h:
                if row["unit"] == "hours": h -= int(row["sm"] or 0)
                elif row["unit"] == "overtime": o -= int(row["sm"] or 0)
            d -= r_d

            # Возвращаем результат (часы переводим в целые числа для интерфейса)
            return {
                "hours": max(0, h // 60), 
                "overtime": max(0, o // 60), 
                "days": max(0, d)
            }

        # 2. СЛУЧАЙ: ТЕКУЩИЙ ГОД
        # Просто берем итог текущего месяца
        summ = compute_month_summary(self.active_db, self._selected_employee_id, self.current_year, self.current_month)
        return {
            "hours": max(0, (summ["start_hours"] + summ["acc_hours"]) // 60),
            "overtime": max(0, (summ["start_overtime"] + summ["acc_overtime"]) // 60),
            "days": max(0, summ["start_days"] + summ["acc_days"])
        }

    @Slot(str, str, str, str)
    def saveMoneyCompList(self, comps_json, order_no, order_date_str, comment):
        if not self.active_db or self._selected_employee_id == 0: return
        try:
            order_date = d_parse(order_date_str)
            comps = json.loads(comps_json)
            with self.active_db.transaction():
                for c in comps:
                    unit, amount = c["unit"], int(c["amount"])
                    is_prev = c.get("usePrevYear", False)
                    
                    if is_prev:
                        # ПРОВЕРКА ЭТАЛОНА
                        balances = self.getAvailableBalances(self.current_year - 1)
                        max_val = balances[unit]
                        if amount > max_val:
                            raise Exception(f"Ошибка: Недостаточно средств в остатках прошлого года. Доступно: {max_val}")

                    # Сохранение
                    if unit in ("hours", "overtime"):
                        self.active_db.add_compensation_money(self._selected_employee_id, unit, amount * 60, None, order_no, order_date, comment)
                    else:
                        self.active_db.add_compensation_money(self._selected_employee_id, unit, None, amount, order_no, order_date, comment)
                    
                    if is_prev:
                        last_id = self.active_db.conn.execute("SELECT last_insert_rowid()").fetchone()[0]
                        self.active_db.conn.execute("UPDATE compensation SET event_date='1900-01-01' WHERE id=?", (last_id,))
                
                # Общая проверка на минусы в течение года
                is_valid, err = validate_non_negative_over_year(self.active_db, self._selected_employee_id, self.current_year)
                if not is_valid: raise Exception(err)
                
            self.refresh_calendar(); self.refresh_employees(); self.refresh_yearly_panorama(); self.loadMoneyComps()
            self.showToast.emit("Приказ сохранен", "success")
        except Exception as e:
            self.showToast.emit(str(e), "error")

    @Slot(str)
    def loadDayDetails(self, date_str):
        if not self.active_db or self._selected_employee_id == 0: return
        
        d0 = d_parse(date_str)
        d_iso_str = d_iso(d0)
        s_dt = datetime.combine(d0, datetime.min.time())
        e_dt = s_dt + timedelta(days=1)
        
        # 1. Загружаем дежурства
        duties = self.active_db.list_duties_for_period(self._selected_employee_id, s_dt, e_dt)
        breaks_map = self.active_db.breaks_for_duty_ids([int(r["id"]) for r in duties])
        
        formatted_duties = []
        for d in duties:
            d_id = int(d["id"])
            b_list = breaks_map.get(d_id, [])
            b_formatted = [{"start": bs.hour*60+bs.minute, "end": be.hour*60+be.minute} for bs, be in b_list]
                
            formatted_duties.append({
                "id": d_id,
                "start": fmt_dt_iso(d["start_dt"])[-5:],
                "end": fmt_dt_iso(d["end_dt"])[-5:],
                "comment": d["comment"] or "",
                "is_shift": bool(int(d["is_shift"] if "is_shift" in d.keys() and d["is_shift"] is not None else 0)),
                "breaks": b_formatted
            })
        self._day_duties = formatted_duties
        self.dayDetailsChanged.emit()

        # 2. Загружаем компенсации (учитывая прошлый год)
        comps = self.active_db.conn.execute("""
            SELECT * FROM compensation WHERE employee_id=? AND method<>'money' AND (
                (event_date = ?) OR 
                (order_date = ?) OR
                (id IN (SELECT compensation_id FROM comp_day_off_date WHERE employee_id=? AND day_off_date = ?))
            )
        """, (self._selected_employee_id, d_iso_str, d_iso_str, self._selected_employee_id, d_iso_str)).fetchall()
        
        formatted_comps = []
        for c in comps:
            unit = c["unit"]
            is_prev = (c["event_date"] == "1900-01-01")
            suffix = " (за прошлый год)" if is_prev else ""

            if unit == "hours":
                type_str = "Ночные часы"
                amount_str = fmt_minutes_ru_words(int(c["amount_minutes"] or 0))
                raw = int(c["amount_minutes"] or 0)
            elif unit == "overtime":
                type_str = "Сверх нормы"
                amount_str = fmt_minutes_ru_words(int(c["amount_minutes"] or 0))
                raw = int(c["amount_minutes"] or 0)
            else:
                type_str = "Дни"
                amount_str = "1 день"
                raw = 1

            formatted_comps.append({
                "id": int(c["id"]),
                "type": type_str + suffix,
                "unit": unit,
                "raw_amount": raw,
                "amount": amount_str,
                "comment": c["comment"] or ""
            })
        self._day_comps = formatted_comps
        self.dayCompsChanged.emit()

    @Slot(int, str)
    def deleteDuty(self, duty_id, current_day_str):
        if not self.active_db: return
        try:
            self.active_db.begin()
            self.active_db.delete_duty(duty_id)
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_calendar()
            self.refresh_yearly_panorama()
            self.refresh_pulse()            
            self.loadDayDetails(current_day_str)
            self.showToast.emit("Дежурство удалено. Отменить: Ctrl+Z", "success")
            
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Не удалось удалить дежурство: {e}", "error")

    @Slot(int, str)
    def deleteCompensation(self, comp_id, current_day_str):
        if not self.active_db: return
        try:
            self.active_db.begin()
            
            # Узнаем, что это за компенсация (дни или часы)
            comp = self.active_db.conn.execute("SELECT unit FROM compensation WHERE id=?", (comp_id,)).fetchone()
            
            if comp and comp["unit"] == "days":
                # Если это дни, получаем все даты этого приказа
                dates_str = self.active_db.get_comp_dates(comp_id)
                if len(dates_str) > 1:
                    # Если дней несколько, оставляем все, КРОМЕ того, который сейчас удаляем
                    from utils import d_parse # На всякий случай импортируем здесь
                    new_dates = [d_parse(x) for x in dates_str if x != current_day_str]
                    self.active_db.replace_comp_dayoff_dates(comp_id, self._selected_employee_id, new_dates)
                else:
                    # Если остался последний день, удаляем приказ целиком
                    self.active_db.delete_compensation(comp_id)
            else:
                # Если это часы - просто удаляем целиком
                self.active_db.delete_compensation(comp_id)

            self.active_db.conn.execute("COMMIT;")

            self.refresh_calendar()
            self.refresh_yearly_panorama()
            self.refresh_pulse()            
            self.loadDayDetails(current_day_str)
            self.showToast.emit("Компенсация удалена. Отменить: Ctrl+Z", "success")

        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Не удалось удалить компенсацию: {e}", "error")

    @Slot(str)
    def clearDayDuties(self, date_str):
        """Удаляет все дежурства в выбранном дне (по клику в меню календаря)"""
        if not self.active_db or self._selected_employee_id == 0: return
        try:
            d0 = d_parse(date_str)
            eid = self._selected_employee_id
            s_dt = datetime.combine(d0, datetime.min.time())
            e_dt = s_dt + timedelta(days=1)
            
            self.active_db.begin()
            duties = self.active_db.list_duties_for_period(eid, s_dt, e_dt)
            
            # Если удалять нечего - выходим (взрыва не будет)
            if not duties:
                self.active_db.conn.execute("ROLLBACK;")
                return
                
            # ...
            for d in duties:
                self.active_db.delete_duty(int(d["id"]))
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_calendar()
            self.refresh_yearly_panorama()
            self.refresh_pulse()            
            self.showToast.emit(f"Дежурства удалены. Отменить: Ctrl+Z", "success")
            # ...
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка удаления: {e}", "error")

    @Slot(str)
    def clearDayCompensations(self, date_str):
        """Удаляет все компенсации в выбранном дне (учитывая прошлый год)"""
        if not self.active_db or self._selected_employee_id == 0: return
        try:
            d_iso0 = d_iso(d_parse(date_str))
            eid = self._selected_employee_id
            
            self.active_db.begin()
            # Ищем ID всех компенсаций, которые затрагивают этот день
            comps = self.active_db.conn.execute("""
                SELECT id, unit, event_date FROM compensation 
                WHERE employee_id=? AND method<>'money' AND (
                    (event_date=?) OR 
                    (order_date=?) OR 
                    (unit='days' AND method='day_off' AND id IN (SELECT compensation_id FROM comp_day_off_date WHERE employee_id=? AND day_off_date=?))
                )
            """, (eid, d_iso0, d_iso0, eid, d_iso0)).fetchall()
            
            if not comps:
                self.active_db.conn.execute("ROLLBACK;")
                return
            
            for c in comps:
                cid = int(c["id"])
                if c["unit"] == "days":
                    dates = self.active_db.get_comp_dates(cid)
                    if len(dates) > 1:
                        # Если дней много, удаляем только один этот день
                        self.active_db.replace_comp_dayoff_dates(cid, eid, [d_parse(x) for x in dates if x != d_iso0])
                    else:
                        # Если день один - удаляем всю запись
                        self.active_db.delete_compensation(cid)
                else:
                    # Часы удаляем целиком
                    self.active_db.delete_compensation(cid)
                    
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_calendar()
            self.refresh_yearly_panorama()
            self.refresh_pulse()            
            self.showToast.emit(f"Компенсации удалены. Отменить: Ctrl+Z", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка удаления: {e}", "error")

    def refresh_groups(self):
        if not self.active_db: return
        formatted_groups = [{"id": 0, "name": "Все", "icon": "Все"}]
        for g in self.active_db.list_groups():
            clean_name = g["name"]
            for ch in "№-.,_()[]{}": clean_name = clean_name.replace(ch, " ")
            icon_text = "".join(p if p.isdigit() else p[0].upper() for p in clean_name.split() if p)[:4] or "?"
            formatted_groups.append({"id": int(g["id"]), "name": g["name"], "icon": icon_text})
        self._group_list = formatted_groups
        self.groupListChanged.emit()

    def refresh_employees(self):
        if not self.active_db: return
        
        # Получаем список из базы (поиск передаем пустой, так как искать будем сами)
        emps = self.active_db.list_employees_for_month(self.current_year, self.current_month, self.show_active_only, "")
        
        group_rows = self.active_db.list_groups()
        group_names = {g["id"]: g["name"] for g in group_rows}
        
        formatted_emps = []
        current_header_gid = -999 

        # Разбиваем то, что ввел пользователь, на отдельные слова и переводим в нижний регистр
        # Например: "Иванов инж" -> ["иванов", "инж"]
        search_words = self.search_text.strip().lower().split()

        for e in emps:
            gid = e["group_id"]
            
            # Фильтр по выбранной вкладке группы слева
            if self.current_group_id is not None and gid != self.current_group_id: 
                continue

            # ==========================================
            # УМНЫЙ ПОИСК НА PYTHON
            # ==========================================
            if search_words:
                # Склеиваем все данные сотрудника в одну длинную строку
                full_info = f"{e['last_name']} {e['first_name']} {e['middle_name'] or ''} {e['rank'] or ''} {e['position'] or ''}".lower()
                
                # Проверяем, есть ли КАЖДОЕ слово из поиска в этой длинной строке
                is_match = True
                for word in search_words:
                    if word not in full_info:
                        is_match = False
                        break
                
                # Если хотя бы одно слово не нашли - пропускаем сотрудника
                if not is_match:
                    continue 

            # ЕСЛИ МЫ ВО ВКЛАДКЕ "ВСЕ", ВСТАВЛЯЕМ ЗАГОЛОВКИ ГРУПП
            if self.current_group_id is None and gid != current_header_gid:
                g_name = group_names.get(gid, "Сотрудники без группы")
                formatted_emps.append({
                    "id": -1, 
                    "is_header": True,
                    "name": g_name,
                    "subtitle": "", "is_active": True, "has_overtime": False,
                    "shift_minutes": 0, "norm_minutes": 0, "last_name": "", "first_name": "", "middle_name": "", "rank": "", "position": "", "start_month": ""
                })
                current_header_gid = gid

            # Собираем красивый текст для карточки
            fio = f"{e['last_name']} {e['first_name']} {e['middle_name'] or ''}".strip()
            sub = " — ".join([x for x in [(e["rank"] or "").strip(), (e["position"] or "").strip()] if x])
            active = is_emp_active_in_month(e, self.current_year, self.current_month)
            
            status = ""
            inactive_reason = "" # Специальная переменная для тултипа
            
            if not active:
                current_m_str = f"{self.current_year:04d}-{self.current_month:02d}"
                # 1. Если месяц приема еще не наступил
                if e["start_month"] > current_m_str:
                    sy, sm = e["start_month"].split("-")
                    months = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"]
                    m_name = months[int(sm) - 1]
                    inactive_reason = f"Будет принят: {m_name} {sy} г."
                # 2. Если уже уволен/переведен
                elif e["end_date"]:
                    if e["end_reason"] == "transfer": 
                        status = f"(переведен {fmt_date_iso(e['end_date'])})"
                        inactive_reason = f"Переведен {fmt_date_iso(e['end_date'])}"
                    elif e["end_reason"] == "dismissal": 
                        status = f"(уволен {fmt_date_iso(e['end_date'])})"
                        inactive_reason = f"Уволен {fmt_date_iso(e['end_date'])}"
                    else: 
                        status = f"(неактивен с {fmt_date_iso(e['end_date'])})"
                        inactive_reason = f"Неактивен с {fmt_date_iso(e['end_date'])}"

            sub2 = (sub + " " + status).strip() if status else sub
            
            is_shift = is_employee_shift(self.active_db, int(e["id"]))
            has_overtime = is_shift
            shift_minutes = 0
            norm_minutes = 0
            
            if is_shift:
                summ = compute_month_summary(self.active_db, int(e["id"]), self.current_year, self.current_month)
                shift_minutes = summ["shift_minutes"]
                norm_minutes = summ["norm_minutes"]
                
            # Замени весь этот кусок внутри цикла for e in emps:
            formatted_emps.append({
                "id": int(e["id"]), 
                "is_header": False, 
                "name": fio, 
                "subtitle": sub2, 
                "is_active": active, 
                "inactive_reason": inactive_reason, 
                "has_overtime": has_overtime,
                "shift_minutes": shift_minutes,
                "norm_minutes": norm_minutes,
                "last_name": e["last_name"],
                "first_name": e["first_name"],
                "middle_name": e["middle_name"] or "",
                "rank": e["rank"] or "",
                "position": e["position"] or "",
                "start_month": e["start_month"],
                "opening_minutes": int(e["opening_minutes"] or 0),
                "opening_overtime": int(e["opening_overtime_minutes"] or 0),
                "opening_days": int(e["opening_days"] or 0),
                "prev_opening_minutes": int(e["prev_opening_minutes"] or 0),
                "prev_opening_overtime": int(e["prev_opening_overtime_minutes"] or 0),
                "prev_opening_days": int(e["prev_opening_days"] or 0)
            })
            
        self._employee_list = formatted_emps
        self.employeeListChanged.emit() 

    def refresh_pulse(self):
        """Проверяет каждый из 12 месяцев: есть ли там данные для 'зеленой точки'"""
        self._month_pulse = {}
        if not self.active_db or self._selected_employee_id == 0:
            self.monthPulseChanged.emit()
            return

        y_str = str(self.current_year)
        eid = self._selected_employee_id

        # Ищем дежурства
        d_rows = self.active_db.conn.execute("SELECT DISTINCT substr(start_dt, 6, 2) as m FROM duty WHERE employee_id=? AND substr(start_dt, 1, 4)=?", (eid, y_str)).fetchall()
        # Ищем статусы (больничные/командировки)
        s_rows = self.active_db.conn.execute("SELECT DISTINCT substr(date, 6, 2) as m FROM employee_day_status WHERE employee_id=? AND substr(date, 1, 4)=?", (eid, y_str)).fetchall()
        # Ищем компенсации (выходные)
        c1_rows = self.active_db.conn.execute("SELECT DISTINCT substr(event_date, 6, 2) as m FROM compensation WHERE employee_id=? AND method<>'money' AND event_date IS NOT NULL AND substr(event_date, 1, 4)=?", (eid, y_str)).fetchall()
        c2_rows = self.active_db.conn.execute("SELECT DISTINCT substr(day_off_date, 6, 2) as m FROM comp_day_off_date WHERE employee_id=? AND substr(day_off_date, 1, 4)=?", (eid, y_str)).fetchall()

        # Собираем всё вместе
        for rows in (d_rows, s_rows, c1_rows, c2_rows):
            for r in rows:
                if r["m"]:
                    try:
                        self._month_pulse[int(r["m"])] = True
                    except: pass

        self.monthPulseChanged.emit()

    def refresh_calendar(self):
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdatescalendar(self.current_year, self.current_month)
        
        if not month_days or not self.active_db:
            return

        grid_start = month_days[0][0]
        grid_end = month_days[-1][-1]
        
        work_map = self.active_db.get_calendar_month(d_iso(grid_start), d_iso(grid_end))
        holidays_set = self.active_db.get_holidays_month(d_iso(grid_start), d_iso(grid_end))
        pre_holidays_set = self.active_db.get_pre_holidays_month(d_iso(grid_start), d_iso(grid_end))
        
        duty_map = {}
        comp_set = set()
        status_map = {}
        
        if self._selected_employee_id > 0:
            eid = self._selected_employee_id
            status_map = self.active_db.get_statuses_for_period(eid, d_iso(grid_start), d_iso(grid_end))
            
            # ЗАГРУЗКА КОМПЕНСАЦИЙ ДЛЯ СЕТКИ
            comps = self.active_db.list_compensations_for_period(eid, d_iso(grid_start), d_iso(grid_end))
            for c in comps:
                if c["unit"] in ("hours", "overtime"):
                    # Если есть визуальная дата (order_date), берем её, иначе event_date
                    display_d = c["order_date"] if c["order_date"] else c["event_date"]
                    if display_d: comp_set.add(display_d)
                else:
                    # Для дней берем все даты из связанной таблицы
                    for cd in self.active_db.get_comp_dates(int(c["id"])):
                        comp_set.add(cd)
            
            start_dt = datetime.combine(grid_start, datetime.min.time())
            end_dt = datetime.combine(grid_end + timedelta(days=1), datetime.min.time())
            duties = self.active_db.list_duties_for_period(eid, start_dt, end_dt)
            breaks_map = self.active_db.breaks_for_duty_ids([int(d["id"]) for d in duties])
            
            for d in duties:
                did = int(d["id"])
                is_shift = bool(int(d["is_shift"] if "is_shift" in d.keys() and d["is_shift"] is not None else 0))
                s0 = max(dt_parse(d["start_dt"]), start_dt)
                e0 = min(dt_parse(d["end_dt"]), end_dt)
                if s0 >= e0: continue
                
                parts = subtract_intervals((s0, e0), breaks_map.get(did, []))
                for s, e in parts:
                    cur = s.date()
                    last = (e - timedelta(seconds=1)).date() if e > s else s.date()
                    while cur <= last:
                        day_start = datetime.combine(cur, datetime.min.time())
                        day_end = day_start + timedelta(days=1)
                        inter = intersect(s, e, day_start, day_end)
                        if inter:
                            t_str = f"{inter[0].strftime('%H:%M')}-{inter[1].strftime('%H:%M')}"
                            duty_map.setdefault(d_iso(cur), []).append({"id": did, "text": t_str, "is_shift": is_shift})
                        cur += timedelta(days=1)

        grid_data = []
        for week in month_days:
            for d in week:
                d_str = d_iso(d)
                is_working = work_map.get(d, d.weekday() < 5)
                is_holiday = d in holidays_set
                grid_data.append({
                    "date_str": d_str,
                    "day_number": d.day,
                    "is_current_month": d.month == self.current_month,
                    "is_weekend": (not is_working) and (not is_holiday),
                    "is_holiday": is_holiday,
                    "is_pre_holiday": d in pre_holidays_set,
                    "status": status_map.get(d, ""),
                    "has_comp": d_str in comp_set,
                    "duties": duty_map.get(d_str, [])
                })
                
        self._calendar_days = grid_data
        self.calendarDaysChanged.emit()
        self.currentPeriodChanged.emit()

        if self._selected_employee_id > 0:
            summ = compute_month_summary(self.active_db, self._selected_employee_id, self.current_year, self.current_month)

            if summ["norm_minutes"] > 0:
                self._selected_employee_ratio = max(0, summ["shift_minutes"]) / summ["norm_minutes"]
            else:
                self._selected_employee_ratio = 0.0
            self.selectedEmployeeRatioChanged.emit() 

            try:
                from export import TemplateExporter
                money_txt = TemplateExporter._money_comp_text_for_month(self.active_db, self._selected_employee_id, self.current_year, self.current_month)
            except Exception:
                money_txt = "—"

            # Умная функция для скобочек в остатках
            # Умная функция для скобочек в остатках (НА НАЧАЛО / НА КОНЕЦ)
            def fmt_dual(val, prev_val, is_days=False):
                main = (f"{val} д." if is_days else fmt_minutes_ru_words(val))
                if prev_val == 0: return main
                # Зеленые скобки для Эталона
                return f"{main} <font color='#4CAF50'>({fmt_minutes_ru_words(prev_val) if not is_days else str(prev_val) + ' д.'})</font>"

            # Умная функция для колонки КОМПЕНСИРОВАНО
            def fmt_comp(real_val, prev_val, is_days=False):
                # Если вообще ничего не списывали - рисуем прочерк
                if real_val == 0 and prev_val == 0: return "—"
                
                # Форматируем основную часть (этот год)
                main = (f"{real_val} д." if is_days else fmt_minutes_ru_words(real_val))
                
                # Если прошлого года нет - отдаем только основную
                if prev_val == 0: return main
                
                # Если этот год 0, а прошлый есть - пишем "0 (за прошлый...)"
                suffix = (f"{prev_val} д." if is_days else fmt_minutes_ru_words(prev_val))
                return f"{main} <font color='#4CAF50'>(за пред. год: {suffix})</font>"

            self._month_summary = {
                "is_shift": summ["is_shift"], 
                "norm_minutes": fmt_minutes_ru_words(summ["norm_minutes"]),
                "shift_minutes": fmt_minutes_ru_words(summ["shift_minutes"]),
                
                "start_hours": fmt_dual(summ["start_hours"], summ["prev_h_start"]),
                "start_overtime": fmt_dual(summ["start_overtime"], summ["prev_o_start"]),
                "start_days": fmt_dual(summ["start_days"], summ["prev_d_start"], True),
                
                "acc_hours": fmt_minutes_ru_words(summ["acc_hours"]),
                "acc_overtime": fmt_minutes_ru_words(summ["acc_overtime"]),
                "acc_days": f"{summ['acc_days']} д.",
                "shift_night": fmt_minutes_ru_words(summ["shift_night"]),
                "shift_holiday": fmt_minutes_ru_words(summ["shift_holiday"]),
                
                # ИСПОЛЬЗУЕМ НОВЫЕ КЛЮЧИ ИЗ LOGIC.PY
                "comp_hours": fmt_comp(summ["comp_h_real"], summ["comp_h_prev"]),
                "comp_overtime": fmt_comp(summ["comp_o_real"], summ["comp_o_prev"]),
                "comp_days": fmt_comp(summ["comp_d_real"], summ["comp_d_prev"], True),
                "comp_money": money_txt,
                
                "end_hours": fmt_dual(summ["end_hours"], summ["prev_h_end"]),
                "end_overtime": fmt_dual(summ["end_overtime"], summ["prev_o_end"]),
                "end_days": fmt_dual(summ["end_days"], summ["prev_d_end"], True),
                
                "is_overtime_negative": summ["end_overtime"] < 0,
                "is_hours_negative": summ["end_hours"] < 0,
                "is_days_negative": summ["end_days"] < 0
            }
        else:
            self._month_summary = {}
        self.monthSummaryChanged.emit()

    def refresh_yearly_panorama(self):
        """Собирает данные за весь год для годовой сетки"""
        if not self.active_db or self._selected_employee_id == 0:
            self._yearly_data = []
            self.yearlyDataChanged.emit()
            return
            
        eid = self._selected_employee_id
        y_start = f"{self.current_year}-01-01"
        y_end = f"{self.current_year}-12-31"

        data_map = {} 
        
        # МАГИЯ: Сначала загружаем ИЗ БАЗЫ весь календарь рабочих/выходных дней на этот год!
        work_map = self.active_db.get_calendar_month(y_start, y_end)

        # 1. Собираем статусы
        sts = self.active_db.conn.execute("SELECT date, status FROM employee_day_status WHERE employee_id=? AND date>=? AND date<=?", (eid, y_start, y_end)).fetchall()
        for r in sts:
            data_map[r["date"]] = {"type": "status", "val": r["status"]}

        # 2. Собираем дежурства
        dts = self.active_db.conn.execute("SELECT start_dt, is_shift FROM duty WHERE employee_id=? AND start_dt>=? AND start_dt<=?", (eid, y_start, y_end+" 23:59")).fetchall()
        for r in dts:
            d_iso_str = r["start_dt"][:10]
            is_shift = int(r["is_shift"] if "is_shift" in r.keys() and r["is_shift"] is not None else 0)
            if d_iso_str not in data_map: 
                data_map[d_iso_str] = {"type": "duty", "val": is_shift}

        # 3. Собираем компенсации (Часы и Дни)
        c1 = self.active_db.conn.execute("SELECT event_date FROM compensation WHERE employee_id=? AND method<>'money' AND event_date IS NOT NULL AND event_date>=? AND event_date<=?", (eid, y_start, y_end)).fetchall()
        for r in c1:
            if r["event_date"] and r["event_date"] not in data_map:
                data_map[r["event_date"]] = {"type": "comp"}
        
        c2 = self.active_db.conn.execute("SELECT day_off_date FROM comp_day_off_date WHERE employee_id=? AND day_off_date>=? AND day_off_date<=?", (eid, y_start, y_end)).fetchall()
        for r in c2:
            if r["day_off_date"] not in data_map:
                data_map[r["day_off_date"]] = {"type": "comp"}

        # 4. Превращаем в матрицу
        matrix = []
        for m in range(1, 13):
            month_row = []
            import calendar as cal_lib
            _, days_in_month = cal_lib.monthrange(self.current_year, m)
            
            for d in range(1, 32):
                if d <= days_in_month:
                    cur_date = date(self.current_year, m, d)
                    d_iso_str = d_iso(cur_date)
                    info = data_map.get(d_iso_str, None)
                    
                    t = ""
                    v = ""
                    if info:
                        t = info["type"]
                        v = str(info.get("val", ""))
                        
                    # МАГИЯ: Смотрим в базу! Если дня нет в базе, по умолчанию Сб/Вс - выходные.
                    # Переменная work_map возвращает True (рабочий) или False (выходной).
                    # is_weekend = это когда НЕ рабочий.
                    is_weekend = not work_map.get(cur_date, cur_date.weekday() < 5)
                        
                    month_row.append({
                        "is_real": True, 
                        "date": d_iso_str,
                        "month": m,
                        "day": d,
                        "is_weekend": is_weekend,
                        "type": t,
                        "val": v
                    })
                else:
                    month_row.append({
                        "is_real": False,
                        "date": "", "month": m, "day": d, "is_weekend": False, "type": "", "val": ""
                    })
                    
            matrix.append(month_row)

        self._yearly_data = matrix
        self.yearlyDataChanged.emit()

        # === СОБИРАЕМ ИТОГИ ЗА ГОД ДЛЯ НИЖНЕЙ ПАНЕЛИ ===
        t_acc_h = t_acc_o = t_acc_d = 0
        t_comp_h = t_comp_o = t_comp_d = 0
        end_h = end_o = end_d = 0
        
        for m in range(1, 13):
            s = compute_month_summary(self.active_db, eid, self.current_year, m)
            t_acc_h += s["acc_hours"]
            t_acc_o += s["acc_overtime"]
            t_acc_d += s["acc_days"]
            t_comp_h += s.get("comp_hours", 0)
            t_comp_o += s.get("comp_overtime", 0)
            t_comp_d += s.get("comp_days", 0)
            if m == 12:
                end_h, end_o, end_d = s["end_hours"], s["end_overtime"], s["end_days"]

        b_days = o_days = k_days = 0
        for date_str, info in data_map.items():
            if info["type"] == "status":
                if info["val"] == "Б": b_days += 1
                elif info["val"] == "О": o_days += 1
                elif info["val"] == "К": k_days += 1

        try:
            from export import TemplateExporter
            mh = self.active_db.conn.execute("SELECT COALESCE(SUM(amount_minutes),0) AS m FROM compensation WHERE employee_id=? AND method='money' AND unit='hours' AND event_date IS NOT NULL AND event_date >= ? AND event_date <= ?", (eid, y_start, y_end)).fetchone()["m"]
            mo = self.active_db.conn.execute("SELECT COALESCE(SUM(amount_minutes),0) AS m FROM compensation WHERE employee_id=? AND method='money' AND unit='overtime' AND event_date IS NOT NULL AND event_date >= ? AND event_date <= ?", (eid, y_start, y_end)).fetchone()["m"]
            md = self.active_db.conn.execute("SELECT COALESCE(SUM(amount_days),0) AS d FROM compensation WHERE employee_id=? AND method='money' AND unit='days' AND event_date IS NOT NULL AND event_date >= ? AND event_date <= ?", (eid, y_start, y_end)).fetchone()["d"]
            
            parts = []
            if mh: parts.append(f"{fmt_minutes_ru_words(mh)} (ноч.)")
            if mo: parts.append(f"{fmt_minutes_ru_words(mo)} (сверх)")
            if md: parts.append(f"{md} дн.")
            money_txt = ", ".join(parts) if parts else "—"
        except:
            money_txt = "—"

        self._year_summary = {
            "acc_hours": fmt_minutes_ru_words(t_acc_h),
            "acc_overtime": fmt_minutes_ru_words(t_acc_o),
            "acc_days": f"{t_acc_d} д.",
            "comp_hours": fmt_minutes_ru_words(t_comp_h),
            "comp_overtime": fmt_minutes_ru_words(t_comp_o),
            "comp_days": f"{t_comp_d} д.",
            "comp_money": money_txt,
            "end_hours": fmt_minutes_ru_words(end_h),
            "end_overtime": fmt_minutes_ru_words(end_o),
            "end_days": f"{end_d} д.",
            "b_days": f"{b_days} дн." if b_days else "—",
            "o_days": f"{o_days} дн." if o_days else "—",
            "k_days": f"{k_days} дн." if k_days else "—",
            "is_hours_negative": end_h < 0,
            "is_overtime_negative": end_o < 0,
            "is_days_negative": end_d < 0
        }
        self.yearSummaryChanged.emit()

    @Slot(str, str)
    def setDayStatus(self, date_str, status_code):
        if not self.active_db or self._selected_employee_id == 0:
            return
        try:
            d0 = d_parse(date_str)
            self.active_db.begin()
            self.active_db.set_day_status(self._selected_employee_id, d0, status_code)
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_calendar()

            self.refresh_yearly_panorama()
            self.refresh_pulse()

        except Exception as e:
            if self.active_db:
                self.active_db.conn.execute("ROLLBACK;")

    @Slot(str, str)
    def setDayType(self, date_str, day_type):
        """Получает дату и тип: 'work', 'weekend' или 'holiday'"""
        if not self.active_db:
            return

        try:
            d0 = d_parse(date_str)
            self.active_db.begin()
            self.active_db.set_calendar_day_type(d0, day_type)
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_calendar()
            #self.refresh_employees()
            self.refresh_yearly_panorama()
            print(f"ПИТОН: День {date_str} стал -> {day_type}")
            
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            print(f"ОШИБКА смены типа дня: {e}")

    @Slot(str, str, str, str, str, str, int, int, int, int, int, int)
    def saveEmployee(self, last_name, first_name, middle_name, rank, position, start_month, open_mins, open_overtime_mins, open_days, prev_mins, prev_overtime, prev_days):
        if not self.active_db: return
        if not last_name.strip() or not first_name.strip():
            print("ПИТОН: Ошибка! Фамилия и имя обязательны.")
            return

        try:
            self.active_db.begin()
            emp_id = self.active_db.add_employee(
                last_name.strip(), first_name.strip(), middle_name.strip(),
                rank.strip(), position.strip(), start_month,
                opening_minutes=open_mins,
                opening_days=open_days,
                opening_overtime=open_overtime_mins,
                prev_opening_minutes=prev_mins,
                prev_opening_overtime=prev_overtime,
                prev_opening_days=prev_days,
                group_id=self.current_group_id
            )
            self.active_db.conn.execute("COMMIT;")
            self.refresh_employees()
            self.showToast.emit("Сотрудник добавлен", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(int, str, str, str, str, str, str, int, int, int, int, int, int)
    def updateEmployee(self, emp_id, last_name, first_name, middle_name, rank, position, start_month, open_mins, open_overtime_mins, open_days, prev_mins, prev_overtime, prev_days):
        if not self.active_db: return
        if not last_name.strip() or not first_name.strip():
            print("ПИТОН: Ошибка! Фамилия и имя обязательны.")
            return

        try:
            self.active_db.begin()
            self.active_db.update_employee(
                emp_id, last_name=last_name.strip(), first_name=first_name.strip(),
                middle_name=middle_name.strip(), rank=rank.strip(), position=position.strip(),
                start_month=start_month, 
                opening_minutes=open_mins, 
                opening_overtime_minutes=open_overtime_mins,
                opening_days=open_days,
                prev_opening_minutes=prev_mins,
                prev_opening_overtime_minutes=prev_overtime,
                prev_opening_days=prev_days
            )
            self.active_db.conn.execute("COMMIT;")
            self.refresh_employees()
            self.refresh_calendar()
            self.showToast.emit("Сотрудник обновлен", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(int)
    def deleteEmployee(self, emp_id):
        if not self.active_db:
            return

        try:
            # МАГИЯ: Открываем блок with.
            # Нам больше не нужно писать COMMIT и ROLLBACK вручную!
            with self.active_db.transaction():
                self.active_db.delete_employee(emp_id)

            # Если мы удалили того сотрудника, который сейчас был выбран - сбрасываем выбор
            if self._selected_employee_id == emp_id:
                self._selected_employee_id = 0
                self.selectedEmployeeChanged.emit()
                self.refresh_calendar()

            self.refresh_employees()
            self.showToast.emit("Сотрудник удалён. Отменить: Ctrl+Z", "success")

        except Exception as e:
            # Ошибка БД автоматически отменит изменения — говорим об этом открыто
            self.showToast.emit(f"Не удалось удалить сотрудника: {e}", "error")

    @Slot(int, str, str)
    def setEmployeeEndDate(self, emp_id, end_date_str, reason):
        """Устанавливает дату и причину окончания работы"""
        if not self.active_db: return
        
        try:
            d0 = d_parse(end_date_str)
            
            self.active_db.begin()
            # update_employee мы уже писали в database.py, она всё сделает!
            self.active_db.update_employee(
                emp_id, 
                end_date=d_iso(d0), 
                end_reason=reason
            )
            self.active_db.conn.execute("COMMIT;")
            
            print(f"ПИТОН: Сотрудник {emp_id} получил статус '{reason}' с {end_date_str}!")
            self.refresh_employees()
            
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            print(f"ОШИБКА установки статуса увольнения: {e}")

    @Slot(int)
    def clearEmployeeEndDate(self, emp_id):
        """Снимает статус увольнения/перевода"""
        if not self.active_db: return
        try:
            self.active_db.begin()
            self.active_db.update_employee(
                emp_id, 
                end_date=None, 
                end_reason=None
            )
            self.active_db.conn.execute("COMMIT;")
            print(f"ПИТОН: Статус увольнения с {emp_id} снят!")
            self.refresh_employees()
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            print(f"ОШИБКА отмены увольнения: {e}")

    @Slot()
    def loadMoneyComps(self):
        if not self.active_db or self._selected_employee_id == 0: return
        
        # Ищем по order_date (визуальной дате приказа)
        current_month_str = f"{self.current_year:04d}-{self.current_month:02d}"
        
        rows = self.active_db.conn.execute(
            """
            SELECT id, unit, order_no, order_date, event_date, amount_minutes, amount_days, comment 
            FROM compensation 
            WHERE employee_id=? 
              AND method='money' 
              AND substr(order_date, 1, 7) = ?
            ORDER BY order_date, id
            """,
            (self._selected_employee_id, current_month_str)
        ).fetchall()
        
        res = []
        for r in rows:
            u = r["unit"]
            # Определяем, был ли это "Прошлый год" по технической дате 1900
            is_prev = (r["event_date"] and r["event_date"].startswith("1900"))
            
            suffix = " (прошлый год)" if is_prev else ""
            
            if u == "hours": 
                typ, amt = "Часы (ночные)" + suffix, fmt_minutes_ru_words(int(r["amount_minutes"] or 0))
                raw_amt = int(r["amount_minutes"] or 0) // 60
            elif u == "overtime": 
                typ, amt = "Сверх нормы" + suffix, fmt_minutes_ru_words(int(r["amount_minutes"] or 0))
                raw_amt = int(r["amount_minutes"] or 0) // 60
            else: 
                typ, amt = "Дни" + suffix, f"{int(r['amount_days'] or 0)} дн."
                raw_amt = int(r["amount_days"] or 0)
            
            res.append({
                "id": int(r["id"]),
                "date": fmt_date_iso(r["order_date"]),
                "type": typ,
                "unit": u,
                "raw_amount": raw_amt,
                "amount": amt,
                "order_no": r["order_no"] or "",
                "comment": r["comment"] or ""
            })
        self._money_comps = res
        self.moneyCompsChanged.emit()

    @Slot(int)
    def deleteMoneyComp(self, comp_id):
        if not self.active_db: return
        try:
            with self.active_db.transaction():
                self.active_db.delete_compensation(comp_id)
            self.refresh_calendar(); self.refresh_employees(); self.refresh_yearly_panorama(); self.loadMoneyComps()
            self.showToast.emit("Выплата удалена", "success")
        except Exception as e:
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot()
    def loadDepartmentData(self):
        """Читает настройки отдела из базы и отправляет в QML"""
        if not self.active_db: return
        try:
            r = self.active_db.get_department_settings()
            if r:
                self._department_data = {
                    "department_name": r["department_name"] or "",
                    "resp_position": r["resp_position"] or "",
                    "resp_rank": r["resp_rank"] or "",
                    "resp_last_name": r["resp_last_name"] or "",
                    "resp_first_name": r["resp_first_name"] or "",
                    "resp_middle_name": r["resp_middle_name"] or ""
                }
                self.departmentDataChanged.emit()
        except Exception as e:
            print(f"ОШИБКА загрузки настроек отдела: {e}")

    @Slot(str, str, str, str, str, str)
    def saveDepartmentData(self, dept_name, pos, rank, last, first, mid):
        """Сохраняет новые настройки отдела в базу"""
        if not self.active_db: return
        try:
            self.active_db.begin()
            self.active_db.update_department_settings(
                department_name=dept_name.strip() or "Отдел",
                resp_position=pos.strip() or None,
                resp_rank=rank.strip() or None,
                resp_last_name=last.strip() or None,
                resp_first_name=first.strip() or None,
                resp_middle_name=mid.strip() or None
            )
            self.active_db.conn.execute("COMMIT;")
            self._active_department_name = dept_name.strip() or "Отдел"
            self.activeDepartmentNameChanged.emit()            
            
            # Название отдела могло измениться, поэтому обновляем список баз
            self.load_databases() 
            self.showToast.emit("Настройки отдела сохранены", "success")
            
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка сохранения: {e}", "error")

    @Slot(str)
    def removeDatabaseFromList(self, path):
        """Удаляет путь из файла config.json (умное сравнение путей)"""
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                
                # Защита: нельзя удалить ту базу, в которой мы сейчас работаем
                if self.active_db and str(Path(self.active_db.path).resolve()) == str(Path(path).resolve()):
                    self.showToast.emit("Нельзя удалить активную базу!", "error")
                    return
                
                # Получаем абсолютный "эталонный" путь того, что хотим удалить
                target_path = str(Path(path).resolve())
                new_paths = []
                
                # Перебираем все сохраненные пути в конфиге
                for p in data.get("db_paths", []):
                    # Превращаем путь из конфига в абсолютный для честного сравнения
                    p_obj = Path(p) if Path(p).is_absolute() else (self.app_dir / p).resolve()
                    
                    if str(p_obj) != target_path:
                        new_paths.append(p) # Если не совпадает, оставляем в списке (сохраняя старый формат)
                        
                data["db_paths"] = new_paths
                
                # Записываем изменения в файл
                self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
                
                # Перезагружаем список в интерфейсе
                self.load_databases()
                self.showToast.emit("База убрана из списка", "success")
                
            except Exception as e:
                self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(str)
    def createNewDatabase(self, name):
        """Создает новый файл .sqlite в стандартной папке и открывает его"""
        if not name.strip():
            self.showToast.emit("Имя базы не может быть пустым", "error")
            return
            
        try:
            # Создаем папку databases рядом с конфигом
            # Создаем папку (или берем пользовательскую)
            db_dir = self._get_db_dir() # <--- ИЗМЕНИТЬ ЗДЕСЬ
            db_dir.mkdir(parents=True, exist_ok=True)
            
            # Очищаем имя от плохих символов
            safe_name = "".join(c for c in name if c.isalnum() or c in " _-").strip()
            if not safe_name: safe_name = "database"
            
            # Ищем уникальное имя файла
            db_path = db_dir / f"{safe_name}.sqlite"
            i = 2
            while db_path.exists():
                db_path = db_dir / f"{safe_name}_{i}.sqlite"
                i += 1
                
            # Создаем пустую базу и пишем в нее имя отдела
            new_db = DB(str(db_path))
            new_db.update_department_settings(department_name=name.strip())
            new_db.conn.commit()
            new_db.close()
            
            # Добавляем в конфиг и загружаем список
            self.add_to_config(str(db_path))
            self.load_databases()
            
            # СРАЗУ открываем эту базу
            self.openDatabase(str(db_path))
            self.showToast.emit(f"База '{name}' успешно создана!", "success")
            
        except Exception as e:
            self.showToast.emit(f"Ошибка создания: {e}", "error")

    @Slot(str)
    def importDatabaseCopy(self, file_url):
        """Создает копию выбранной базы в нашей системной папке и подключает её"""
        if not file_url: return
        path = QUrl(file_url).toLocalFile() if file_url.startswith("file://") else file_url

        # Проверяем файл ДО копирования: не всякий выбранный файл — наша база
        ok, err = self._validate_db_file(path)
        if not ok:
            self.showToast.emit(f"Не удалось импортировать: {err}", "error")
            return

        try:
            db_dir = self._get_db_dir()
            db_dir.mkdir(parents=True, exist_ok=True)
            
            # Читаем имя отдела из импортируемой базы, чтобы красиво назвать файл
            temp_db = DB(path)
            dept_name = temp_db.get_department_name()
            temp_db.close()
            
            safe_name = "".join(c for c in dept_name if c.isalnum() or c in " _-").strip() or "Imported_DB"
            new_path = db_dir / f"{safe_name}_{uuid.uuid4().hex[:6]}.sqlite"
            
            import shutil
            shutil.copy2(path, str(new_path))
            
            self.add_to_config(str(new_path))
            self.load_databases()
            self.showToast.emit("База успешно импортирована (скопирована)", "success")
            
        except Exception as e:
            self.showToast.emit(f"Ошибка импорта: {e}", "error")

    @Slot(str, str)
    def exportDatabaseCopy(self, current_db_path, folder_url):
        """Копирует базу из нашей системы в папку, которую выбрал пользователь"""
        if not current_db_path or not folder_url: return
        
        target_folder = QUrl(folder_url).toLocalFile() if folder_url.startswith("file://") else folder_url
        
        try:
            temp_db = DB(current_db_path)
            dept_name = temp_db.get_department_name()
            temp_db.close()

            safe_name = "".join(c for c in dept_name if c.isalnum() or c in " _-").strip() or "Database"
            export_path = Path(target_folder) / f"{safe_name}.sqlite"

            # ЗАЩИТА ОТ ПЕРЕЗАПИСИ: если файл с таким именем уже есть в папке,
            # подбираем свободное имя ("База_2.sqlite", "База_3.sqlite"...),
            # а не затираем чужую копию молча.
            i = 2
            while export_path.exists():
                export_path = Path(target_folder) / f"{safe_name}_{i}.sqlite"
                i += 1

            import shutil
            shutil.copy2(current_db_path, str(export_path))

            self.showToast.emit(f"Копия базы сохранена: {export_path.name}", "success")
        except Exception as e:
            self.showToast.emit(f"Ошибка экспорта: {e}", "error")

    @Slot()
    def loadHotkeys(self):
        """Читает список горячих клавиш из конфига и отдает в QML"""
        if not self.config_path.exists(): return
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
            ui_cfg = data.get("ui", {})
            self._hotkeys_list = ui_cfg.get("custom_shortcuts", [])
            self.hotkeysListChanged.emit()
        except Exception as e:
            print(f"ОШИБКА загрузки хоткеев: {e}")

    @Slot(str)
    def saveHotkeys(self, hotkeys_json):
        """Сохраняет массив горячих клавиш обратно в конфиг"""
        try:
            data = {"db_paths": [], "last_db_path": None, "ui": {}}
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            
            ui_cfg = data.get("ui", {})
            ui_cfg["custom_shortcuts"] = json.loads(hotkeys_json)
            data["ui"] = ui_cfg
            
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            self.loadHotkeys() # Обновляем список в интерфейсе
            self.showToast.emit("Горячие клавиши сохранены", "success")
        except Exception as e:
            self.showToast.emit(f"Ошибка сохранения хоткеев: {e}", "error")

    @Slot(str, str)
    def executeHotkey(self, key_sequence, target_date_str):
        if not self.active_db or self._selected_employee_id == 0: return
        
        target_action = None
        for sc in self._hotkeys_list:
            if sc["key"] == key_sequence:
                target_action = sc
                break
                
        if not target_action: return 

        # --- ИСПРАВЛЕНИЕ: Делаем проверку ДО открытия транзакции! ---
        if target_action["type"] == "duty":
            is_shift = target_action.get("duty_shift", False)
            if is_shift and not self.isSelectedEmployeeShift:
                self.showToast.emit("Ошибка: Хоткей сменного графика не применим к пятидневщику", "error")
                return
        # -----------------------------------------------------------

        try:
            d0 = d_parse(target_date_str)
            self.active_db.begin() # <--- Теперь открываем базу безопасно
            
            # --- ДОБАВЛЕНА ОБРАБОТКА СТАТУСОВ ---
            if target_action["type"] == "status":
                self.active_db.set_day_status(self._selected_employee_id, d0, target_action["status_val"])
                
            elif target_action["type"] == "duty":
                s_time = parse_hhmm(target_action["duty_start"])
                e_time = parse_hhmm(target_action["duty_end"])
                
                start_dt = datetime.combine(d0, s_time)
                end_dt = datetime.combine(d0, e_time)
                if end_dt <= start_dt: end_dt += timedelta(days=1)
                
                is_shift = target_action.get("duty_shift", False)
                
                duty_id = self.active_db.add_duty(self._selected_employee_id, start_dt, end_dt, f"Hotkey: {key_sequence}", is_shift)
                
                breaks = target_action.get("duty_breaks", [])
                if breaks:
                    breaks_list = []
                    for b in breaks:
                        bs_time = parse_hhmm(b["start"])
                        be_time = parse_hhmm(b["end"])
                        
                        bs_dt = datetime.combine(d0, bs_time)
                        if bs_dt < start_dt: bs_dt += timedelta(days=1)
                        be_dt = datetime.combine(bs_dt.date(), be_time)
                        if be_dt <= bs_dt: be_dt += timedelta(days=1)
                        
                        breaks_list.append((bs_dt, be_dt))
                    self.active_db.replace_duty_breaks(duty_id, breaks_list)
                    
            elif target_action["type"] == "comp":
                unit = target_action["comp_unit"]
                amount = int(target_action["comp_amount"])
                if unit in ("hours", "overtime"):
                    self.active_db.add_compensation_hours_dayoff(self._selected_employee_id, d0, amount * 60, f"Hotkey: {key_sequence}", unit=unit)
                else:
                    dates_list = [d0 + timedelta(days=i) for i in range(amount)]
                    self.active_db.add_compensation_days_dayoff(self._selected_employee_id, dates_list, f"Hotkey: {key_sequence}")
            
            is_valid, err = validate_non_negative_over_year(self.active_db, self._selected_employee_id, d0.year)
            if not is_valid:
                self.active_db.conn.execute("ROLLBACK;")
                self.showToast.emit(err, "error")
                return

            self.active_db.conn.execute("COMMIT;")
            self.refresh_calendar()
            self.refresh_employees()
            self.refresh_yearly_panorama()
            self.refresh_pulse()
            self.showToast.emit(f"Применено: {key_sequence}", "success")
            
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка хоткея: {e}", "error")

    @Slot(str, bool)
    def createGroup(self, name, is_shift):
        if not self.active_db or not name.strip(): return
        try:
            self.active_db.begin()
            self.active_db.add_group(name, is_shift)
            self.active_db.conn.execute("COMMIT;")
            self.refresh_groups()
            self.showToast.emit("Группа создана", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(int)
    def deleteGroup(self, group_id):
        if not self.active_db: return
        try:
            self.active_db.begin()
            self.active_db.delete_group(group_id)
            self.active_db.conn.execute("COMMIT;")
            if self.current_group_id == group_id:
                self.current_group_id = None
            self.refresh_groups()
            self.refresh_employees()
            self.showToast.emit("Группа удалена. Отменить: Ctrl+Z", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(int, result=int)
    def getGroupEmployeeCount(self, group_id):
        """Сколько сотрудников сейчас в группе (для предупреждения при удалении группы)"""
        if not self.active_db: return 0
        try:
            row = self.active_db.conn.execute(
                "SELECT COUNT(*) FROM employee WHERE group_id=?", (group_id,)
            ).fetchone()
            return int(row[0]) if row else 0
        except Exception:
            return 0

    @Slot(int, int)
    def moveEmployeeToGroup(self, emp_id, group_id):
        if not self.active_db: return
        try:
            target_gid = None if group_id == 0 else group_id
            self.active_db.begin()
            self.active_db.set_employee_group(emp_id, target_gid)
            self.active_db.conn.execute("COMMIT;")
            
            # 1. Обновляем список слева
            self.refresh_employees()
            
            # Если перетащили того сотрудника, который сейчас открыт на экране
            if emp_id == self._selected_employee_id:
                # 2. Обновляем свойство для скрытия галочки
                self.selectedEmployeeChanged.emit()
                
                # 3. ПОЛНОСТЬЮ ПЕРЕСЧИТЫВАЕМ И ПЕРЕРИСОВЫВАЕМ ЕГО ДАННЫЕ
                self.refresh_calendar()
                self.refresh_yearly_panorama()
                self.refresh_pulse()
            
            self.showToast.emit("Сотрудник перемещен", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(int, int, str)
    def officialTransferEmployee(self, emp_id, group_id, date_str):
        """Shift + Перетаскивание (Официальный перевод). Сохраняет историю."""
        if not self.active_db: return
        try:
            d0 = d_parse(date_str)
            target_gid = None if group_id == 0 else group_id
            self.active_db.begin()
            
            # 1. Проверяем, есть ли уже история у этого человека
            count = self.active_db.conn.execute("SELECT COUNT(*) as c FROM employee_transfer WHERE employee_id=?", (emp_id,)).fetchone()["c"]
            if count == 0:
                # МАГИЯ: Берем реальный месяц приема на работу сотрудника!
                emp_data = self.active_db.get_employee(emp_id)
                current_group = emp_data["group_id"]
                # Превращаем "2022-05" в "2022-05-01"
                start_date = f"{emp_data['start_month']}-01" if emp_data["start_month"] else "2000-01-01"
                
                self.active_db.conn.execute("INSERT INTO employee_transfer (employee_id, transfer_date, group_id) VALUES (?, ?, ?)", (emp_id, start_date, current_group))
                
            # 2. Делаем запись о НОВОМ переводе
            self.active_db.conn.execute("INSERT INTO employee_transfer (employee_id, transfer_date, group_id) VALUES (?, ?, ?)", (emp_id, d_iso(d0), target_gid))
            
            # 3. Меняем текущую группу
            self.active_db.set_employee_group(emp_id, target_gid)
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_employees()
            if emp_id == self._selected_employee_id:
                self.selectedEmployeeChanged.emit()
                self.refresh_calendar()
                self.refresh_yearly_panorama()
                self.refresh_pulse()
            self.showToast.emit(f"Официальный перевод с {date_str} сохранен", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка перевода: {e}", "error")

    @Slot(str)
    def exportToExcel(self, file_url):
        """Запускает сборку Excel в невидимом фоновом потоке"""
        if not self.active_db:
            self.showToast.emit("Сначала откройте базу данных", "error")
            return
            
        try:
            out_path = QUrl(file_url).toLocalFile() if file_url.startswith("file://") else file_url
            if not out_path.endswith(".xlsx"):
                out_path += ".xlsx"
                
            template_path = EXCEL_TEMPLATE_PATH
            
            # ИСПРАВЛЕНИЕ ЗДЕСЬ:
            # Если это обычный путь (DEV), проверяем через Path. Если ресурс (PROD) - пропускаем проверку.
            if not template_path.startswith(":/") and not Path(template_path).exists():
                self.showToast.emit(f"Шаблон не найден: {Path(template_path).name}", "error")
                return
                
            # Показываем уведомление, что процесс пошел (чтобы пользователь не кликал дважды)
            self.showToast.emit("Создание Excel-файла... (Фоновый режим)", "success")
            
            # МАГИЯ: Запускаем Работягу
            # Сохраняем в self.export_thread, чтобы сборщик мусора Питона не убил поток раньше времени
            self.export_thread = ExportWorker(
                db_path=self.active_db.path,
                year=self.current_year,
                month=self.current_month,
                template_path=template_path, # Убрали str(), так как это уже строка
                out_path=out_path
            )
            
            # Подключаем "уши", чтобы слушать ответ от Работяги
            self.export_thread.finished_signal.connect(self.on_export_finished)
            self.export_thread.start() # Поехали!
            
        except Exception as e:
            self.showToast.emit(f"Ошибка запуска экспорта: {e}", "error")

    def on_export_finished(self, success, message):
        """Срабатывает автоматически, когда поток закончил работу"""
        if success:
            self.showToast.emit(message, "success")
        else:
            self.showToast.emit(message, "error")

    @Slot(str, int, str, str, str, str, bool)
    def quickPrint(self, printer_name, copies, page_from, page_to, orientation, paper_size, collate):
        """Бесшовная скрытая печать в фоновом потоке"""
        if not self.active_db:
            self.showToast.emit("Сначала откройте базу данных", "error")
            return
            
        template_path = EXCEL_TEMPLATE_PATH
        
        # ИСПРАВЛЕНИЕ ЗДЕСЬ:
        if not template_path.startswith(":/") and not Path(template_path).exists():
            self.showToast.emit(f"Шаблон не найден: {Path(template_path).name}", "error")
            return

        # Показываем уведомление, что процесс пошел
        self.showToast.emit("Формирование документа... (Фоновый режим)", "success")

        try:
            # Запускаем Работягу
            self.print_thread = PrintWorker(
                db_path=self.active_db.path,
                year=self.current_year,
                month=self.current_month,
                printer_name=printer_name,
                copies=copies,
                page_from=page_from,
                page_to=page_to,
                orientation=orientation,
                paper_size=paper_size,
                collate=collate
            )
            self.print_thread.finished_signal.connect(self.on_print_finished)
            self.print_thread.start()
        except Exception as e:
            self.showToast.emit(f"Не удалось запустить печать: {e}", "error")

    def on_print_finished(self, success, message):
        """Срабатывает автоматически, когда фоновая печать завершена"""
        if success:
            self.showToast.emit(message, "success")
        else:
            self.showToast.emit(message, "error")

    @Slot(str, str)
    def handleClipboard(self, action_type, date_str):
        """Обрабатывает копирование, вырезание, вставку и удаление (Клавиша Del)"""
        if not self.active_db or self._selected_employee_id == 0:
            return

        try:
            d0 = d_parse(date_str)
            eid = self._selected_employee_id
            d_iso0 = d_iso(d0)
            
            s_dt = datetime.combine(d0, datetime.min.time())
            e_dt = s_dt + timedelta(days=1)
        except Exception as e:
            self.showToast.emit(f"Ошибка даты: {e}", "error")
            return

        # ========================
        # УДАЛЕНИЕ (Клавиша Delete)
        # ========================
        if action_type == "delete":
            # 1. Ищем дежурства
            d_rows = self.active_db.list_duties_for_period(eid, s_dt, e_dt)
            
            # 2. Ищем компенсации (УЧИТЫВАЕМ ПРОШЛЫЙ ГОД через order_date)
            c_rows = self.active_db.conn.execute("""
                SELECT id, unit FROM compensation WHERE employee_id=? AND method<>'money' AND (
                    (event_date = ?) OR 
                    (order_date = ?) OR
                    (id IN (SELECT compensation_id FROM comp_day_off_date WHERE employee_id=? AND day_off_date = ?))
                )
            """, (eid, d_iso0, d_iso0, eid, d_iso0)).fetchall()
            
            # 3. Ищем статусы
            status_row = self.active_db.conn.execute("SELECT status FROM employee_day_status WHERE employee_id=? AND date=?", (eid, d_iso0)).fetchone()
            
            if not d_rows and not c_rows and not status_row:
                self.showToast.emit("В этот день ничего нет", "error")
                return

            try:
                self.active_db.begin()
                
                # Удаляем статус
                if status_row:
                    self.active_db.conn.execute("DELETE FROM employee_day_status WHERE employee_id=? AND date=?", (eid, d_iso0))
                
                # Удаляем дежурства
                for d in d_rows:
                    self.active_db.delete_duty(int(d["id"]))
                    
                # Удаляем компенсации
                for c in c_rows:
                    cid = int(c["id"])
                    if c["unit"] == "days":
                        dates = self.active_db.get_comp_dates(cid)
                        if len(dates) > 1:
                            self.active_db.replace_comp_dayoff_dates(cid, eid, [d_parse(x) for x in dates if x != d_iso0])
                        else:
                            self.active_db.delete_compensation(cid)
                    else:
                        # Часы (включая прошлый год) удаляем полностью
                        self.active_db.delete_compensation(cid)
                        
                self.active_db.conn.execute("COMMIT;")
                self.refresh_calendar()
                self.refresh_yearly_panorama()
                self.refresh_pulse()
                self.showToast.emit(f"Очищено: {date_str}", "success")
                
            except Exception as e:
                self.active_db.conn.execute("ROLLBACK;")
                self.showToast.emit(f"Ошибка удаления: {e}", "error")
            return

        # ========================
        # КОПИРОВАНИЕ / ВЫРЕЗАНИЕ
        # ========================
        if action_type in ("copy", "cut"):
            try:
                duties_data = []
                duties = self.active_db.list_duties_for_period(eid, s_dt, e_dt)
                breaks_map = self.active_db.breaks_for_duty_ids([int(r["id"]) for r in duties])
                
                for r in duties:
                    if dt_parse(r["start_dt"]).date() == d0:
                        is_sh = bool(int(r["is_shift"] if r["is_shift"] is not None else 0))
                        duties_data.append({
                            "start": dt_parse(r["start_dt"]), 
                            "end": dt_parse(r["end_dt"]),
                            "comment": r["comment"] or "", 
                            "is_shift": is_sh,
                            "breaks": breaks_map.get(int(r["id"]), [])
                        })
                
                # Копируем компенсации (тоже учитываем order_date)
                comps = self.active_db.conn.execute("""
                    SELECT * FROM compensation WHERE employee_id=? AND method<>'money' AND (
                        (event_date = ?) OR (order_date = ?) OR
                        (id IN (SELECT compensation_id FROM comp_day_off_date WHERE employee_id=? AND day_off_date = ?))
                    )
                """, (eid, d_iso0, d_iso0, eid, d_iso0)).fetchall()
                
                comps_data = [{"unit": r["unit"], "amount_minutes": int(r["amount_minutes"] or 0), "comment": r["comment"] or "", "event_date": r["event_date"]} for r in comps]
                
                if not duties_data and not comps_data:
                    self.showToast.emit("День пуст", "error")
                    return
                    
                self._clipboard = {"src_date": d0, "duties": duties_data, "comps": comps_data}
                
                if action_type == "cut":
                    self.handleClipboard("delete", date_str) 
                    self.showToast.emit("Вырезано", "success")
                else:
                    self.showToast.emit("Скопировано", "success")
            except Exception as e:
                self.showToast.emit(f"Ошибка копирования: {e}", "error")
            return

        # ========================
        # ВСТАВКА
        # ========================
        if action_type == "paste":
            if not self._clipboard:
                self.showToast.emit("Буфер обмена пуст!", "error")
                return
            try:
                src_date = self._clipboard["src_date"]
                duties = self._clipboard["duties"]
                comps = self._clipboard["comps"]
                delta = timedelta(days=(d0 - src_date).days)
                
                self.active_db.begin()
                for d in duties:
                    new_id = self.active_db.add_duty(eid, d["start"] + delta, d["end"] + delta, d["comment"], d["is_shift"])
                    self.active_db.replace_duty_breaks(new_id, [(b[0]+delta, b[1]+delta) for b in d["breaks"]])
                    
                for c in comps:
                    is_prev = (c["event_date"] == "1900-01-01")
                    # Передаем признак прошлого года при вставке
                    if c["unit"] == "hours":
                        self.saveCompensation(d_iso0, "hours", str(c["amount_minutes"]), c["comment"], is_prev)
                    else:
                        self.saveCompensation(d_iso0, "days", "1", c["comment"], is_prev)
                        
                self.active_db.conn.execute("COMMIT;")
                self.refresh_calendar()
                self.showToast.emit(f"Вставлено в {date_str}", "success")
            except Exception as e:
                self.active_db.conn.execute("ROLLBACK;")
                self.showToast.emit(f"Ошибка вставки: {e}", "error")

    @Slot()
    def undoAction(self):
        """Отменяет последнее действие (Ctrl+Z)"""
        if not self.active_db: return
        if self.active_db.undo():
            self.refresh_calendar()
            self.refresh_employees()
            self.refresh_groups()
            self.showToast.emit("Действие отменено (Ctrl+Z)", "success")
        else:
            self.showToast.emit("Больше нечего отменять", "error")

    @Slot()
    def redoAction(self):
        """Возвращает отмененное действие (Ctrl+Y)"""
        if not self.active_db: return
        if self.active_db.redo():
            self.refresh_calendar()
            self.refresh_employees()
            self.refresh_groups()
            self.showToast.emit("Действие возвращено (Ctrl+Y)", "success")
        else:
            self.showToast.emit("Больше нечего возвращать", "error")

    @Slot(int)  # <--- ВОТ ЭТА МАГИЧЕСКАЯ СТРОЧКА!
    def setYear(self, new_year):
        """Меняет год и перерисовывает всё"""
        if self.current_year != new_year:
            self.current_year = new_year
            self.refresh_calendar()
            self.refresh_employees()
            self.refresh_yearly_panorama()
            self.refresh_pulse()

    @Slot()
    def generateYearList(self):
        """Генерирует список годов (например, +-5 лет от текущего)"""
        current = date.today().year
        # Создаем список от 2020 до 2030
        years = [y for y in range(current - 5, current + 6)]
        self._year_list = years
        self.yearListChanged.emit()

    @Slot(int, int)
    def reorderGroups(self, source_id, target_id):
        if not self.active_db: return
        groups = self.active_db.list_groups()
        ordered_ids = [g["id"] for g in groups]
        
        if source_id in ordered_ids and target_id in ordered_ids:
            idx_s = ordered_ids.index(source_id)
            idx_t = ordered_ids.index(target_id)
            if idx_s == idx_t: return
            
            # Удаляем старый элемент
            item = ordered_ids.pop(idx_s)
            # Находим новый индекс цели (так как список сдвинулся)
            new_idx_t = ordered_ids.index(target_id)
            # Вставляем строго ПЕРЕД целью
            ordered_ids.insert(new_idx_t, item)
            
            self.active_db.begin()
            self.active_db.update_group_orders(ordered_ids)
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_groups()
            self.refresh_employees()
            self.showToast.emit("Порядок групп изменен", "success")

    @Slot(int, int)
    def reorderEmployees(self, source_id, target_id):
        if not self.active_db: return
        current_ids = [e["id"] for e in self._employee_list if not e["is_header"]]
        
        if source_id in current_ids and target_id in current_ids:
            idx_s = current_ids.index(source_id)
            idx_t = current_ids.index(target_id)
            if idx_s == idx_t: return
            
            # Удаляем старый элемент
            item = current_ids.pop(idx_s)
            # Находим новый индекс цели
            new_idx_t = current_ids.index(target_id)
            # Вставляем строго ПЕРЕД целью
            current_ids.insert(new_idx_t, item)
            
            self.active_db.begin()
            self.active_db.update_employee_orders(current_ids)
            self.active_db.conn.execute("COMMIT;")
            
            self.refresh_employees()
            self.showToast.emit("Порядок сотрудников изменен", "success")

    @Slot(str)
    def setTimeInputMode(self, mode):
        """Сохраняет выбранный стиль ввода времени ('slider' или 'tumbler')"""
        if self._time_input_mode == mode: return
        self._time_input_mode = mode
        
        try:
            data = {"db_paths": [], "last_db_path": None, "ui": {}}
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            
            ui_cfg = data.get("ui", {})
            ui_cfg["time_input_mode"] = mode
            data["ui"] = ui_cfg
            
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            self.timeInputModeChanged.emit() # Кричим QML, чтобы он мгновенно перестроился
            self.showToast.emit("Внешний вид обновлен", "success")
        except Exception as e:
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(int)
    def loadTransferHistory(self, emp_id):
        if not self.active_db: return
        
        rows = self.active_db.get_employee_transfers(emp_id)
        emp_data = self.active_db.get_employee(emp_id)
        
        # Вычисляем возможные даты "невидимых" якорей
        anchor_date1 = "2000-01-01"
        anchor_date2 = f"{emp_data['start_month']}-01" if emp_data["start_month"] else "2000-01-01"
        
        res = []
        for r in rows:
            # МАГИЯ: Если это системный якорь - просто пропускаем его, не отдаем в интерфейс
            if r["transfer_date"] in (anchor_date1, anchor_date2):
                continue
                
            res.append({
                "id": int(r["id"]),
                "date": fmt_date_iso(r["transfer_date"]), 
                "raw_date": r["transfer_date"],           
                "group_id": r["group_id"] or 0,
                "group_name": r["group_name"] or "Без группы"
            })
            
        self._transfer_history = res
        self.employeeTransferHistoryChanged.emit()

    @Slot(int, int, str, int)
    def saveTransferRecord(self, emp_id, record_id, date_str, group_id):
        if not self.active_db: return
        try:
            d0 = d_parse(date_str)
            target_gid = None if group_id == 0 else group_id
            
            self.active_db.begin()
            if record_id == 0: # Добавление нового
                self.active_db.conn.execute("INSERT INTO employee_transfer (employee_id, transfer_date, group_id) VALUES (?, ?, ?)", (emp_id, d_iso(d0), target_gid))
            else: # Обновление старого
                self.active_db.conn.execute("UPDATE employee_transfer SET transfer_date=?, group_id=? WHERE id=?", (d_iso(d0), target_gid, record_id))
            
            self.active_db.sync_employee_current_group(emp_id)
            self.active_db.conn.execute("COMMIT;")
            
            # Обновляем всё на экране
            self.loadTransferHistory(emp_id)
            self.refresh_employees()
            if emp_id == self._selected_employee_id:
                self.selectedEmployeeChanged.emit()
                self.refresh_calendar()
                self.refresh_yearly_panorama()
                
            self.showToast.emit("Запись сохранена", "success")
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка сохранения: {e}", "error")

    @Slot(int, int)
    def deleteTransferRecord(self, emp_id, record_id):
        if not self.active_db: return
        try:
            self.active_db.begin()
            self.active_db.conn.execute("DELETE FROM employee_transfer WHERE id=?", (record_id,))
            self.active_db.sync_employee_current_group(emp_id)
            self.active_db.conn.execute("COMMIT;")
            
            self.loadTransferHistory(emp_id)
            self.refresh_employees()
            if emp_id == self._selected_employee_id:
                self.selectedEmployeeChanged.emit()
                self.refresh_calendar()
                self.refresh_yearly_panorama()
        except Exception as e:
            self.active_db.conn.execute("ROLLBACK;")
            self.showToast.emit(f"Ошибка удаления: {e}", "error")

    @Slot(int, str, str, str, bool, str, str)
    def updateDuty(self, duty_id, date_str, start_time_str, end_time_str, is_shift, comment, breaks_json):
        if not self.active_db or self._selected_employee_id == 0: return
        try:
            d0 = d_parse(date_str)
            s_time = parse_hhmm(start_time_str)
            e_time = parse_hhmm(end_time_str)
            start_dt = datetime.combine(d0, s_time)
            end_dt = datetime.combine(d0, e_time)
            if end_dt <= start_dt: end_dt += timedelta(days=1)

            overlaps = self.active_db.find_overlapping_duties(self._selected_employee_id, start_dt, end_dt, exclude_duty_id=duty_id)
            if overlaps:
                self.showToast.emit("Ошибка: Пересечение с другим дежурством", "error")
                return

            with self.active_db.transaction():
                self.active_db.conn.execute(
                    "UPDATE duty SET start_dt=?, end_dt=?, comment=?, is_shift=? WHERE id=?",
                    (dt_iso(start_dt), dt_iso(end_dt), comment or None, int(is_shift), duty_id)
                )
                if breaks_json:
                    breaks_data = json.loads(breaks_json)
                    breaks_list = []
                    for b in breaks_data:
                        bs_time = time(int(b["start_h"]), int(b["start_m"]))
                        be_time = time(int(b["end_h"]), int(b["end_m"]))
                        bs_dt = datetime.combine(d0, bs_time)
                        if bs_dt < start_dt: bs_dt += timedelta(days=1)
                        be_dt = datetime.combine(bs_dt.date(), be_time)
                        if be_dt <= bs_dt: be_dt += timedelta(days=1)
                        breaks_list.append((bs_dt, be_dt))
                    self.active_db.replace_duty_breaks(duty_id, breaks_list)
                    
            self.refresh_calendar()
            self.loadDayDetails(date_str)
            self.showToast.emit("Дежурство обновлено", "success")
        except Exception as e:
            self.showToast.emit(f"Ошибка обновления: {e}", "error")

    @Slot(int, str, str, str, str, bool)
    def updateCompensation(self, comp_id, date_str, comp_type, amount_str, comment, use_prev_year):
        if not self.active_db or self._selected_employee_id == 0: return
        try:
            logic_date = "1900-01-01" if use_prev_year else date_str

            with self.active_db.transaction():
                if comp_type in ("hours", "overtime"):
                    self.active_db.conn.execute(
                        "UPDATE compensation SET unit=?, amount_minutes=?, amount_days=NULL, comment=?, event_date=? WHERE id=?",
                        (comp_type, int(amount_str), comment or None, logic_date, comp_id)
                    )
                else:
                    self.active_db.conn.execute(
                        "UPDATE compensation SET unit='days', amount_minutes=NULL, amount_days=1, comment=?, event_date=? WHERE id=?",
                        (comment or None, logic_date, comp_id)
                    )

                is_valid, err = validate_non_negative_over_year(self.active_db, self._selected_employee_id, self.current_year)
                if not is_valid:
                    raise Exception(err)

            self.refresh_calendar()
            self.loadDayDetails(date_str)
        except Exception as e:
            self.showToast.emit(str(e), "error")

    @Slot(int, str, str, str, str)
    def updateMoneyComp(self, comp_id, unit, amount_str, order_no, comment):
        if not self.active_db or self._selected_employee_id == 0: return
        try:
            amount = int(amount_str)
            with self.active_db.transaction():
                # Обновляем
                if unit in ("hours", "overtime"):
                    self.active_db.conn.execute("UPDATE compensation SET amount_minutes=?, order_no=?, comment=? WHERE id=?", (amount * 60, order_no, comment or None, comp_id))
                else:
                    self.active_db.conn.execute("UPDATE compensation SET amount_days=?, order_no=?, comment=? WHERE id=?", (amount, order_no, comment or None, comp_id))
                
                # Получаем год события для проверки
                row = self.active_db.conn.execute("SELECT event_date FROM compensation WHERE id=?", (comp_id,)).fetchone()
                event_date = d_parse(row["event_date"])
                
                is_valid, error_msg = validate_non_negative_over_year(self.active_db, self._selected_employee_id, event_date.year)
                if not is_valid: raise Exception(error_msg)
                
            self.refresh_calendar()
            self.refresh_yearly_panorama()
            self.loadMoneyComps()
            self.showToast.emit("Приказ обновлен", "success")
        except Exception as e:
            self.showToast.emit(f"Ошибка: {e}", "error")

    @Slot(str)
    def openDbFolder(self, path):
        """Открывает папку с базой данных в проводнике Windows"""
        import platform, os, subprocess
        folder = str(Path(path).parent.resolve())
        if not Path(folder).exists():
            self.showToast.emit("Папка не найдена", "error")
            return
            
        try:
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            self.showToast.emit(f"Ошибка открытия папки: {e}", "error")

    @Slot(str, str, result="QVariant")
    def getShiftDatesForPeriod(self, start_date_str: str, end_date_str: str):
        """
        Анализирует паттерн сменщика и возвращает даты его рабочих дней
        в заданном периоде. Вызывается из DayEventDialog.
        """
        if not self.active_db or self._selected_employee_id == 0:
            return {"dates": [], "error": "Нет активного сотрудника", "confidence": 0.0}
        
        try:
            from shift_pattern import get_shift_work_dates_for_period
            from utils import d_parse
            
            start = d_parse(start_date_str)
            end = d_parse(end_date_str)
            
            if end < start:
                return {"dates": [], "error": "Дата конца раньше даты начала", "confidence": 0.0}
            
            result = get_shift_work_dates_for_period(
                self.active_db,
                self._selected_employee_id,
                start,
                end
            )
            
            return result
            
        except Exception as e:
            return {"dates": [], "error": str(e), "confidence": 0.0}

    @Slot(str)
    def changeDbDirectory(self, folder_url):
        """Меняет стандартную папку и переносит туда все базы"""
        import shutil
        target_dir = QUrl(folder_url).toLocalFile() if folder_url.startswith("file://") else folder_url
        target_path = Path(target_dir).resolve()
        
        try:
            target_path.mkdir(parents=True, exist_ok=True)
            data = {"db_paths": [], "ui": {}}
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            
            new_paths = []
            active_path_resolved = str(Path(self.active_db.path).resolve()) if self.active_db else None
            
            # Временно закрываем активную базу, чтобы Windows разрешил переместить файл!
            if self.active_db:
                self.active_db.close()
            
            for p in data.get("db_paths", []):
                p_obj = Path(p) if Path(p).is_absolute() else (self.app_dir / p).resolve()
                if p_obj.exists():
                    new_file_path = target_path / p_obj.name
                    
                    if p_obj != new_file_path:
                        # Защита от перезаписи файлов с одинаковым именем
                        i = 2
                        while new_file_path.exists() and new_file_path.resolve() != p_obj.resolve():
                            new_file_path = target_path / f"{p_obj.stem}_{i}{p_obj.suffix}"
                            i += 1
                        shutil.move(str(p_obj), str(new_file_path))
                    
                    new_paths.append(str(new_file_path))
                else:
                    new_paths.append(p)
                    
            data["db_paths"] = new_paths
            
            if "ui" not in data: data["ui"] = {}
            data["ui"]["default_db_dir"] = str(target_path)
            
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            
            # Переоткрываем базу на новом месте в тихом режиме
            if active_path_resolved:
                for np in new_paths:
                    if Path(np).name == Path(active_path_resolved).name:
                        self.active_db = DB(np)
                        break
                        
            self.load_databases()
            self.showToast.emit(f"Все базы перенесены в {target_path.name}", "success")
            
        except Exception as e:
            self.showToast.emit(f"Ошибка переноса: {e}", "error")

    # ====================================================
    # ОБНОВЛЕНИЕ ПРОГРАММЫ (офлайн, как кнопка в Telegram)
    # data/ не трогаем: базы, хоткеи и тема остаются на месте.
    # ====================================================

    def _read_qrc_version(self) -> str:
        try:
            from PySide6.QtCore import QFile, QIODevice
            f = QFile(":/components/AppTheme.qml")
            if f.open(QIODevice.ReadOnly):
                text = bytes(f.readAll()).decode("utf-8", errors="replace")
                f.close()
                return app_update._read_version_from_text(text)
        except Exception:
            return ""
        return ""

    def _init_updates(self):
        qrc_ver = self._read_qrc_version()
        if qrc_ver:
            self._app_version = qrc_ver
        if self._app_version:
            try:
                root = app_update.install_root(self.app_dir)
                app_update.write_version_json(root / "version.json", self._app_version)
                if root != Path(self.app_dir):
                    app_update.write_version_json(Path(self.app_dir) / "version.json", self._app_version)
            except Exception:
                pass
        # Если прошлый раз уже обновились — подчистить хвосты
        staged = app_update.staged_info(self.app_dir)
        if staged:
            ver = staged.get("version") or ""
            if ver and not app_update.is_newer(ver, self._app_version):
                app_update.cleanup_pending(self.app_dir)
        self.scanForUpdates()

    @Property(str, notify=appVersionChanged)
    def appVersion(self):
        return self._app_version or ""

    @Property(bool, notify=updateReadyChanged)
    def updateReady(self):
        return self._update_ready

    @Property(bool, notify=updateBusyChanged)
    def updateBusy(self):
        return self._update_busy

    @Property(str, notify=updateVersionChanged)
    def updateVersion(self):
        return self._update_version

    @Property(str, notify=updateStatusTextChanged)
    def updateStatusText(self):
        return self._update_status

    @Property(int, notify=updateReadyChanged)
    def updateChromeExtra(self):
        return 52 if (self._update_ready or self._update_busy) else 0

    def _set_update_busy(self, busy, text=""):
        self._update_busy = bool(busy)
        self._update_status = text or ""
        self.updateBusyChanged.emit()
        self.updateStatusTextChanged.emit()
        self.updateReadyChanged.emit()  # высота полоски тоже зависит от busy

    def _set_update_ready(self, ready, version=""):
        self._update_ready = bool(ready)
        if version:
            self._update_version = version
            self.updateVersionChanged.emit()
        self.updateReadyChanged.emit()

    @Slot()
    def scanForUpdates(self):
        """Ищет zip/папку новой версии рядом с программой и на флешках."""
        if self._update_busy:
            return
        try:
            dismissed = ""
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                dismissed = str(data.get("ui", {}).get("dismissed_update_version") or "")
        except Exception:
            dismissed = ""

        staged = app_update.staged_info(self.app_dir)
        if staged and staged.get("version") and app_update.is_newer(staged["version"], self._app_version):
            if staged["version"] != dismissed:
                self._update_source = staged["root"]
                self._set_update_ready(True, staged["version"])
            return

        found = app_update.pick_best_update(self.app_dir, self._app_version)
        if not found:
            return
        if found["version"] and found["version"] == dismissed:
            return
        if found.get("already_staged"):
            self._update_source = found["root"]
            self._set_update_ready(True, found["version"])
            return
        self.prepareUpdateFromPath(found["root"])

    @Slot(str)
    def prepareUpdateFromPath(self, file_url):
        """Готовит обновление из zip или папки. Можно вызвать из диалога настроек."""
        if not file_url or self._update_busy:
            return
        path = QUrl(file_url).toLocalFile() if str(file_url).startswith("file:") else file_url
        if not path:
            path = file_url
        try:
            info = app_update.describe_package(Path(path))
        except Exception as e:
            self.showToast.emit(str(e), "error")
            return

        ver = info.get("version") or ""
        if not ver:
            self.showToast.emit("В архиве нет номера версии — так обновляться нельзя", "error")
            return
        if not app_update.is_newer(ver, self._app_version):
            if ver == self._app_version:
                self.showToast.emit(f"Это та же версия ({ver})", "error")
            else:
                self.showToast.emit(f"Откат запрещён: {ver} старше текущей {self._app_version}", "error")
            return

        self._set_update_busy(True, "Готовим обновление…")
        self.showToast.emit("Готовим обновление. Можно продолжать работу.", "success")
        self.update_thread = UpdateStageWorker(info["root"], str(self.app_dir))
        self.update_thread.finished_signal.connect(self._on_update_staged)
        self.update_thread.start()

    def _on_update_staged(self, ok, message, version):
        self._set_update_busy(False, "")
        if not ok:
            self.showToast.emit(f"Не удалось подготовить обновление: {message}", "error")
            return
        staged = app_update.staged_info(self.app_dir)
        self._update_source = staged["root"] if staged else ""
        self._set_update_ready(True, version or (staged or {}).get("version") or "")
        label = self._update_version or "новой версии"
        self.showToast.emit(f"Готово. Можно обновить до {label}", "success")

    @Slot()
    def applyReadyUpdate(self):
        """Кнопка «Обновить»: переодеваем коробку и перезапускаемся. data/ не трогаем."""
        staged = app_update.staged_info(self.app_dir)
        if not staged:
            self.showToast.emit("Сначала укажите файл или папку новой версии", "error")
            return
        staged_ver = staged.get("version") or ""
        if not staged_ver or not app_update.is_newer(staged_ver, self._app_version):
            self.showToast.emit("Откат на старую версию запрещён", "error")
            return
        dest_root = app_update.install_root(self.app_dir)
        if not (dest_root / "OVERTIMETAB.exe").exists() and not IS_FROZEN:
            self.showToast.emit("Обновление ставится в собранную программу, не из редактора кода", "error")
            return
        try:
            if self.active_db:
                try:
                    self.active_db.close()
                except Exception:
                    pass
                self.active_db = None
            app_update.launch_file_swap(
                Path(staged["root"]),
                dest_root,
                os.getpid(),
            )
            # Помощник ждёт смерти PID. Не прячемся в трей и не ждём таймер —
            # иначе старый процесс живёт, а консоль обновления крутится вечно.
            os._exit(0)
        except Exception as e:
            self.showToast.emit(f"Не удалось запустить обновление: {e}", "error")

    @Slot()
    def dismissUpdate(self):
        """Спрятать кнопку до следующего более нового файла."""
        try:
            data = {"db_paths": [], "last_db_path": None, "ui": {}}
            if self.config_path.exists():
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
            ui_cfg = data.get("ui", {})
            ui_cfg["dismissed_update_version"] = self._update_version or ""
            data["ui"] = ui_cfg
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
        self._set_update_ready(False)

import ctypes # <--- Добавляем системную библиотеку Windows

# Уникальный ID сервера для общения между процессами
APP_ID = "OvertimeTab_v2_SingleInstance_Key"

def main():
    try:
        myappid = 'mycompany.overtimetab.version2'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    os.environ["QT_QUICK_CONTROLS_STYLE"] = "Basic" 
    
    app = QApplication(sys.argv)

    # --- ЛОГИКА ОДНОГО ЭКЗЕМПЛЯРА ---
    socket = QLocalSocket()
    socket.connectToServer(APP_ID)
    
    if socket.waitForConnected(500):
        # Если нашли работающий сервер - пишем ему "проснись" и закрываемся
        socket.write(b"RESTORE")
        socket.waitForBytesWritten(500)
        socket.disconnectFromServer()
        return 
    
    # Если сервера нет - создаем его
    server = QLocalServer()
    server.removeServer(APP_ID) # На всякий случай чистим мусор
    server.listen(APP_ID)
    # --------------------------------

    app_icon = QIcon(APP_ICON_PATH)
    app.setWindowIcon(app_icon)

    tray_icon = QSystemTrayIcon(app_icon, app)
    tray_icon.setToolTip("OVERTIMETAB — табель учёта переработок")
    
    tray_menu = QMenu()
    open_action = QAction("Развернуть", app)
    exit_action = QAction("Закрыть полностью", app)
    
    tray_menu.addAction(open_action)
    tray_menu.addAction(exit_action)
    tray_icon.setContextMenu(tray_menu)
    tray_icon.show()

    my_backend = Backend()
    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("backend", my_backend)
    
    def show_window():
        if engine.rootObjects():
            window = engine.rootObjects()[0]
            window.show()
            window.raise_()
            window.requestActivate()

    # --- ОБРАБОТЧИК СИГНАЛА "ПРОСНИСЬ" ---
    def handle_new_connection():
        new_socket = server.nextPendingConnection()
        if new_socket.waitForReadyRead(500):
            msg = new_socket.readAll().data().decode()
            if msg == "RESTORE":
                show_window()

    server.newConnection.connect(handle_new_connection)
    # -------------------------------------

    open_action.triggered.connect(show_window)
    exit_action.triggered.connect(app.quit) 
    
    def tray_activated(reason):
        # Одинарный клик тоже разворачивает окно — как в Telegram и Discord
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick, QSystemTrayIcon.MiddleClick):
            show_window()
            
    tray_icon.activated.connect(tray_activated)

    engine.load(QML_PATH)
    
    if not engine.rootObjects(): 
        sys.exit(-1)
        
    app.setQuitOnLastWindowClosed(False)

    exit_code = app.exec()
    server.close() # Закрываем сервер при выходе
    del engine
    sys.exit(exit_code)

if __name__ == "__main__":
    apply = app_update.parse_apply_argv(sys.argv)
    if apply:
        try:
            app_update.apply_update_inplace(
                Path(apply["source"]),
                Path(apply["dest"]),
                apply.get("wait_pid") or None,
            )
        except Exception as e:
            print(f"Ошибка применения обновления: {e}")
            sys.exit(1)
        sys.exit(0)
    main()