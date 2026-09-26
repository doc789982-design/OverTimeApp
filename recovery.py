# -*- coding: utf-8 -*-
"""
Аварийный восстановитель OverTimeTab — «BIOS» программы.

ЗАДАЧА: если программа упала (в том числе при запуске и даже при загрузке
Qt), при следующем включении ЭТОТ модуль — первым делом, до всех прочих
импортов — предложит:
  1) обновиться на свежую (уже починенную) версию,
  2) откатиться на предыдущую (снимок делает каждое обновление),
  3) запуститься как есть.

ПРАВИЛА ВЫЖИВАНИЯ (нарушать нельзя):
  * только стандартная библиотека, никаких Qt/win32/print;
  * app_update импортируется ЛЕНИВО и только для ветки «обновиться» —
    если он повреждён, откат и обычный запуск всё равно работают;
  * ни одна функция не имеет права уронить программу: всё в try/except;
  * этот файл после отладки ЗАМОРОЖЕН: правим только при крайней нужде.

Состояние сессии — Documents/OverTimeTab/run_state.json:
  running  — программа запущена (pid). Мёртвый pid = авария или выключение.
  updating — программа сама завершается ради обновления/отката (не авария).
  crashed  — поймана непредвиденная ошибка (sys.excepthook).
  ok       — программа открылась и работает.
Журнал восстановителя — Documents/OverTimeTab/recovery.log.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from pathlib import Path

STATE_NAME = "run_state.json"
LOG_NAME = "recovery.log"
ROLLBACK_DIRNAME = "rollback"
ROLLBACK_META_NAME = "ROLLBACK_META.json"
CAPTION = "OverTimeTab — восстановление"

# Кнопки/иконки MessageBoxW (нативное окно, работает без Qt).
_MB_OK = 0
_MB_YESNO = 4
_MB_YESNOCANCEL = 3
_MB_ICONERROR = 0x10
_MB_ICONWARNING = 0x30
_MB_ICONINFO = 0x40
_MB_TOPMOST = 0x40000 | 0x10000  # TOPMOST | SETFOREGROUND
_ID_OK, _ID_CANCEL, _ID_YES, _ID_NO = 1, 2, 6, 7


# ---------------------------------------------------------------------------
# Базовые пути (никогда не бросают исключений)
# ---------------------------------------------------------------------------

def _home() -> Path:
    try:
        home = (os.environ.get("OVERTIMETAB_RECOVERY_HOME")
                or os.environ.get("USERPROFILE")
                or os.environ.get("HOME"))
        return (Path(home) / "Documents" if home else Path.home() / "Documents") / "OverTimeTab"
    except Exception:
        try:
            return Path.home() / "Documents" / "OverTimeTab"
        except Exception:
            return Path(".") / "OverTimeTab"


def _state_path() -> Path:
    return _home() / STATE_NAME


def _rlog(msg: str) -> None:
    """Журнал восстановителя (с обрезкой, чтобы не рос вечно)."""
    try:
        p = _home() / LOG_NAME
        p.parent.mkdir(parents=True, exist_ok=True)
        if p.exists() and p.stat().st_size > 200_000:
            p.unlink()
        with open(p, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "  " + msg + "\n")
    except Exception:
        pass


def _read_state() -> dict:
    try:
        return json.loads(_state_path().read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_state(**kw) -> None:
    try:
        p = _state_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(".tmp")
        kw.setdefault("pid", os.getpid())
        kw.setdefault("ts", time.strftime("%Y-%m-%d %H:%M:%S"))
        tmp.write_text(json.dumps(kw, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, p)
    except Exception:
        _rlog("не удалось записать состояние: " + traceback.format_exc(limit=2))


def _pid_alive(pid: int) -> bool:
    """Жив ли процесс (ctypes, без запуска tasklist)."""
    pid = int(pid or 0)
    if pid <= 0 or pid == os.getpid():
        return pid > 0
    try:
        if os.name == "nt":
            import ctypes
            k32 = ctypes.windll.kernel32
            h = k32.OpenProcess(0x1000, 0, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
            if not h:
                return False
            try:
                code = ctypes.c_ulong()
                if k32.GetExitCodeProcess(h, ctypes.byref(code)):
                    return code.value == 259  # STILL_ACTIVE
                return True
            finally:
                k32.CloseHandle(h)
        os.kill(pid, 0)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Диалоги (нативные, без Qt)
# ---------------------------------------------------------------------------

def _msg(text: str, buttons: int = _MB_OK, icon: int = _MB_ICONWARNING) -> int:
    """Окно сообщения. Не сумелось — «Отмена» (безопасный путь: запустить)."""
    try:
        import ctypes
        return int(ctypes.windll.user32.MessageBoxW(
            None, text, CAPTION, buttons | icon | _MB_TOPMOST))
    except Exception:
        _rlog("диалог недоступен: " + text[:100].replace("\n", " | "))
        return _ID_CANCEL


def _info_while(text: str, worker: threading.Thread) -> None:
    """Показывает окно «идёт работа» и закрывает его сам, когда worker закончит."""
    tid = threading.main_thread().ident

    def _closer():
        try:
            import ctypes
            user32 = ctypes.windll.user32
            found = []

            def _cb(h, _lp):
                found.append(h)
                user32.PostMessageW(h, 0x0010, 0, 0)  # WM_CLOSE
                return True

            CB = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p)
            worker.join(900)
            for _ in range(60):  # ждём появления окна
                del found[:]
                user32.EnumThreadWindows(tid, CB(_cb), 0)
                if found:
                    break
                if not worker.is_alive():
                    break
                time.sleep(0.25)
        except Exception:
            pass

    threading.Thread(target=_closer, daemon=True).start()
    _msg(text, _MB_OK, _MB_ICONINFO)


# ---------------------------------------------------------------------------
# Ветка «обновиться» (app_update — лениво, повреждение не смертельно)
# ---------------------------------------------------------------------------

def _load_app_update():
    """Модуль обновлений, если он цел. Иначе ветка обновления недоступна."""
    try:
        au = __import__("app_update")
        for attr in ("fetch_update_info", "is_newer", "current_app_version",
                     "current_app_build", "install_root", "resolve_download_url",
                     "download_zip", "stage_package", "launch_file_swap",
                     "DEFAULT_UPDATE_URL", "FALLBACK_UPDATE_URL"):
            if not hasattr(au, attr):
                return None
        return au
    except Exception:
        return None


def _update_urls(au) -> list:
    urls = []
    try:
        cfg = _home() / "config.json"
        if cfg.exists():
            ui = (json.loads(cfg.read_text(encoding="utf-8")) or {}).get("ui", {})
            stored = str((ui or {}).get("update_url") or "").strip()
            if stored:
                urls.append(stored)
    except Exception:
        pass
    urls += [au.DEFAULT_UPDATE_URL, au.FALLBACK_UPDATE_URL]
    seen, out = set(), []
    for u in urls:
        u = (u or "").strip()
        if u and u not in seen:
            seen.add(u)
            out.append(u)
    return out


def _do_update(au, rollback) -> None:
    """Проверить обновления и поставить. Любая неудача — откат или запуск."""
    root = _install_root()
    if au is None or root is None:
        _msg("Проверка обновлений сейчас недоступна.\nЗапускаем программу как есть.",
             _MB_OK, _MB_ICONERROR)
        return
    try:
        cur_ver = au.current_app_version(root)
        cur_bld = au.current_app_build(root)
    except Exception:
        cur_ver, cur_bld = "", 0
    found = None
    for u in _update_urls(au):
        try:
            _rlog("восстановление: проверяем " + u)
            info, _err = au.fetch_update_info(u, timeout=8)
            if info and au.is_newer(str(info.get("version") or ""), cur_ver,
                                    int(info.get("build") or 0), cur_bld):
                found = info
                break
        except Exception:
            continue
    if not found:
        _rlog("восстановление: обновлений не нашлось")
        text = "Обновлений не нашлось (нет сети или у вас последняя версия)."
        if rollback:
            if _msg(text + "\n\nОткатиться на сборку %s?\n\nДА — откатиться\nНЕТ — запустить как есть"
                    % rollback.get("build"), _MB_YESNO, _MB_ICONWARNING) == _ID_YES:
                _do_rollback(rollback)
        else:
            _msg(text + "\nЗапускаем программу как есть.", _MB_OK, _MB_ICONINFO)
        return
    bld = found.get("build") or "?"
    try:
        dl = au.resolve_download_url(str(found.get("url") or ""),
                                     str(found.get("base_url") or ""))
        outcome = {}

        def _download():
            try:
                outcome["zip"] = au.download_zip(
                    dl, root / "update_download",
                    sha256=str(found.get("sha256") or "").strip(),
                    timeout=120)
            except BaseException as e:  # причина сбоя попадёт в окно ошибки
                outcome["err"] = e

        worker = threading.Thread(target=_download, daemon=True)
        worker.start()
        _info_while("Скачиваем обновление (сборка %s).\nЭто окно закроется само — просто подождите." % bld, worker)
        worker.join(60)
        if "err" in outcome:
            raise outcome["err"]
        if "zip" not in outcome:
            raise OSError("скачивание не завершилось")
        local_zip = outcome["zip"]
    except Exception as e:
        _rlog("восстановление: скачивание не удалось: %s %s" % (type(e).__name__, e))
        if rollback and _msg("Не удалось скачать обновление (%s).\n\nОткатиться на сборку %s?\n\nДА — откатиться\nНЕТ — запустить как есть"
                             % (e, rollback.get("build")), _MB_YESNO, _MB_ICONERROR) == _ID_YES:
            _do_rollback(rollback)
        else:
            _msg("Запускаем программу как есть.", _MB_OK, _MB_ICONINFO)
        return
    try:
        _rlog("восстановление: готовим пакет " + str(local_zip))
        au.stage_package(Path(local_zip), root)
        au.launch_file_swap(root / "pending_update", root, os.getpid())
        mark_updating()
        _rlog("восстановление: передаём установщику, выходим")
        os._exit(0)
    except Exception as e:
        _rlog("восстановление: установка не удалась: %s %s" % (type(e).__name__, e))
        if rollback and _msg("Не удалось установить обновление (%s).\n\nОткатиться на сборку %s?\n\nДА — откатиться\nНЕТ — запустить как есть"
                             % (e, rollback.get("build")), _MB_YESNO, _MB_ICONERROR) == _ID_YES:
            _do_rollback(rollback)
        else:
            _msg("Запускаем программу как есть.", _MB_OK, _MB_ICONINFO)


# ---------------------------------------------------------------------------
# Ветка «откатиться» (полностью самостоятельная — не зависит от app_update)
# ---------------------------------------------------------------------------

_VBS_TEMPLATE = r'''Option Explicit
Dim src, dst, pid, exe, sh, fso, logFile, t0, rc, q
q = Chr(34)
src = WScript.Arguments(0)
dst = WScript.Arguments(1)
pid = WScript.Arguments(2)
exe = dst & "\OVERTIMETAB.exe"
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
logFile = sh.ExpandEnvironmentStrings("%TEMP%") & "\overtimetab_rollback.log"
WriteLog "recovery start src=" & src & " dst=" & dst & " pid=" & pid
t0 = Timer
Do While PidAlive(pid)
  If (Timer - t0) > 45 Then
    WriteLog "wait timeout"
    Exit Do
  End If
  WScript.Sleep 400
Loop
WScript.Sleep 800
If Not fso.FileExists(src & "\OVERTIMETAB.exe") Then
  WriteLog "missing source exe"
  WScript.Quit 1
End If
rc = sh.Run("robocopy " & q & src & q & " " & q & dst & q & " /E /XD data pending_update update_download /XF *.sqlite *.sqlite-wal *.sqlite-shm /NFL /NDL /NJH /NJS /NC /NS /NP /R:3 /W:1", 0, True)
WriteLog "robocopy=" & rc
If rc >= 8 Then
  WriteLog "robocopy failed"
  If fso.FileExists(exe) Then sh.Run q & exe & q, 1, False
  WScript.Quit 1
End If
If fso.FileExists(exe) Then
  sh.Run q & exe & q, 1, False
  WriteLog "restarted"
End If
fso.DeleteFile WScript.ScriptFullName, True
WScript.Quit 0

Function PidAlive(p)
  Dim wmi, procs
  On Error Resume Next
  PidAlive = False
  If p = "" Or p = "0" Then Exit Function
  Set wmi = GetObject("winmgmts:\\.\root\cimv2")
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  Set procs = wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId=" & CLng(p))
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  If procs.Count > 0 Then PidAlive = True
End Function

Sub WriteLog(msg)
  Dim ts
  On Error Resume Next
  Set ts = fso.OpenTextFile(logFile, 8, True)
  ts.WriteLine Now & " " & msg
  ts.Close
End Sub
'''


def _swap_vbs(source: Path, dest: Path, pid: int) -> None:
    """Свой помощник отката: ждёт смерти pid, копирует, перезапускает.
    В отличие от установщика обновлений, папку-источник НЕ удаляет."""
    vbs = Path(tempfile.gettempdir()) / ("ot_recovery_%d.vbs" % os.getpid())
    vbs.write_text(_VBS_TEMPLATE, encoding="ascii", errors="replace")
    windir = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    wscript = windir / "System32" / "wscript.exe"
    if not wscript.exists():
        wscript = Path("wscript.exe")
    creation = 0x08000000  # CREATE_NO_WINDOW
    startup = None
    if hasattr(subprocess, "STARTUPINFO"):
        startup = subprocess.STARTUPINFO()
        startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startup.wShowWindow = 0
    subprocess.Popen(
        [str(wscript), "//B", "//Nologo", str(vbs), str(source), str(dest), str(pid)],
        cwd=str(tempfile.gettempdir()), close_fds=True, creationflags=creation,
        startupinfo=startup, stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _do_rollback(rollback) -> None:
    root = _install_root()
    if root is None or not rollback.get("dir"):
        _msg("Сохранённая предыдущая версия не найдена.\nЗапускаем программу как есть.",
             _MB_OK, _MB_ICONERROR)
        return
    try:
        _rlog("восстановление: откат на %s" % rollback.get("dir"))
        _swap_vbs(Path(rollback["dir"]), root, os.getpid())
        mark_updating()
        _rlog("восстановление: откат передан помощнику, выходим")
        os._exit(0)
    except Exception as e:
        _rlog("восстановление: откат не удался: %s %s" % (type(e).__name__, e))
        _msg("Не удалось откатиться (%s).\nЗапускаем программу как есть." % e,
             _MB_OK, _MB_ICONERROR)


def rollback_info() -> dict:
    """Снимок предыдущей версии (его делает каждое обновление)."""
    try:
        d = _home() / ROLLBACK_DIRNAME
        if not d.is_dir():
            return None
        if not ((d / "OVERTIMETAB.exe").exists() or (d / "overtimetab.exe").exists()):
            return None
        meta = {"dir": str(d), "version": "", "build": 0, "display": ""}
        p = d / ROLLBACK_META_NAME
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            meta["version"] = str(data.get("version") or "")
            meta["build"] = int(data.get("build") or 0)
            meta["display"] = str(data.get("display") or "")
        return meta
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Главный вход
# ---------------------------------------------------------------------------

def _install_root():
    try:
        d = Path(sys.executable).resolve().parent
        if (d / "OVERTIMETAB.exe").exists() or (d / "overtimetab.exe").exists():
            return d
        return d  # даже без exe — лучшего места у нас нет
    except Exception:
        return None


def _helper_launch(argv) -> bool:
    """Служебные запуски (обновление из pending, веб-режим) — не вмешиваемся."""
    if "--apply-update" in argv:
        return True
    for a in argv[1:]:
        if a.lower().rstrip(" .,;:!") in ("--web", "-web", "/web"):
            return True
    return False


def guard() -> None:
    """Первый код программы. Читает прошлую сессию и решает: авария или нет."""
    try:
        if _helper_launch(sys.argv):
            return
        if not (getattr(sys, "frozen", False)
                or os.environ.get("OVERTIMETAB_RECOVERY_TEST")):
            return  # запуск из редактора кода — восстановитель молчит
        state = _read_state()
        status = str(state.get("status") or "")
        try:
            pid = int(state.get("pid") or 0)
        except Exception:
            pid = 0
        if status in ("running", "updating", "crashed", "ok") and pid \
                and pid != os.getpid() and _pid_alive(pid):
            return  # другой экземпляр жив — маркер не трогаем
        crashed = status == "crashed" or (
            status == "running" and not _pid_alive(pid))
        if not crashed:
            _write_state(status="running")
            return
        err = str(state.get("err") or "")
        _rlog("авария прошлой сессии (status=%s pid=%s): %s"
              % (status, pid, err[:150]))
        rollback = rollback_info()
        head = "OverTimeTab аварийно завершилась в прошлый раз."
        if err:
            head += "\nПоследняя ошибка: " + err.splitlines()[-1][:160]
        if rollback:
            head += ("\n\nДА — проверить обновления и починить (нужен интернет, до пары минут)\n"
                     "НЕТ — откатиться на предыдущую версию (сборка %s)\n"
                     "ОТМЕНА — запустить как есть" % (rollback.get("build") or "?"))
            choice = _msg(head, _MB_YESNOCANCEL, _MB_ICONWARNING)
            if choice == _ID_YES:
                _do_update(_load_app_update(), rollback)
            elif choice == _ID_NO:
                _do_rollback(rollback)
        else:
            head += ("\n\nДА — проверить обновления и починить (нужен интернет, до пары минут)\n"
                     "НЕТ — запустить как есть")
            if _msg(head, _MB_YESNO, _MB_ICONWARNING) == _ID_YES:
                _do_update(_load_app_update(), None)
        # дошли сюда — пользователь выбрал «запустить как есть»
        _write_state(status="running")
    except Exception:
        _rlog("guard: " + traceback.format_exc(limit=3))


# ---------------------------------------------------------------------------
# Метки состояния (зовёт Main.py; каждая защищена от ошибок)
# ---------------------------------------------------------------------------

def _frozen_like() -> bool:
    return bool(getattr(sys, "frozen", False)
                or os.environ.get("OVERTIMETAB_RECOVERY_TEST"))


def mark_ok() -> None:
    """Программа открылась и работает — прошлую сессию считаем удачной."""
    if _frozen_like():
        _write_state(status="ok")


def mark_crash(details: str = "") -> None:
    """Непредвиденная ошибка: следующее включение предложит восстановление."""
    if _frozen_like():
        _write_state(status="crashed", err=str(details or "")[:2000])


def mark_updating() -> None:
    """Завершаемся ради обновления/отката — это не авария."""
    if _frozen_like():
        _write_state(status="updating")
