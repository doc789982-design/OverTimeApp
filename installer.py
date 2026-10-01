# -*- coding: utf-8 -*-
"""Установка и удаление OVERTIMETAB (мастер работает в Main.py с --setup).

Разделение труда:
  tools/installer_stub.py — крошечная заглушка в релизном exe: распаковывает
    программу во временную папку и запускает с --setup;
  этот модуль — вся логика установки/удаления: папки, копирование,
    ярлыки, запись в «Установка и удаление программ»;
  components/SetupWizard.qml — окна мастера в стиле программы.

Куда ставим (по умолчанию):
  «для меня»    %LOCALAPPDATA%\\Programs\\OVERTIMETAB — без прав администратора,
                автообновление проходит молча;
  «для всех»    %ProgramFiles%\\OVERTIMETAB — классика, но требует UAC.

Обновление на месте использует копирование поверх + штатную чистку
мусора прошлых сборок (app_files.txt, см. app_update.cleanup_stale_files —
она вызывается при первом запуске новой версии).

Пользоватские данные (базы, отчёты) живут в Documents\\OverTimeTab и
установкой/обновлением не затрагиваются. Удаление трогает их только
по явной галке в окне деинсталлятора.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

APP_DIRNAME = "OVERTIMETAB"
LNK_NAME = "OVERTIMETAB.lnk"
UNINSTALL_REG = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\OVERTIMETAB"
EXE_NAME = "OVERTIMETAB.exe"

IS_WINDOWS = sys.platform == "win32"


# ────────────────────────── папки ──────────────────────────

def default_install_dir(per_user: bool = True, env=None) -> Path:
    """Стандартная папка установки. Чистая функция — тестируется везде."""
    env = os.environ if env is None else env
    if per_user:
        base = env.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))
        return Path(base) / "Programs" / APP_DIRNAME
    base = env.get("ProgramFiles", r"C:\Program Files")
    return Path(base) / APP_DIRNAME


def known_install_dirs(env=None) -> list[Path]:
    """Где искать уже установленную копию (обе схемы + подсказка реестра)."""
    env = os.environ if env is None else env
    out = []
    for per_user in (True, False):
        d = default_install_dir(per_user, env)
        if d not in out:
            out.append(d)
    return out


def data_root(home: str = "") -> Path:
    """Documents\\OverTimeTab — базы и настройки (см. батч 210)."""
    base = Path(home) if home else Path.home()
    return base / "Documents" / "OverTimeTab"


def _shell_folder(csidl: int) -> Path:
    """Служебные папки Windows (рабочий стол, меню «Пуск»)."""
    import ctypes
    from ctypes import wintypes
    buf = ctypes.create_unicode_buffer(260)
    ctypes.windll.shell32.SHGetFolderPathW(None, csidl, None, 0, buf)
    return Path(buf.value)


def shortcut_paths(desktop: bool, start_menu: bool, per_user: bool = True) -> list[Path]:
    """Пути .lnk, которые нужно создать (или убрать при деинсталляции)."""
    out = []
    if not IS_WINDOWS:
        return out
    if desktop:
        csidl = 0x0000 if per_user else 0x0019        # DESKTOP / COMMON_DESKTOP
        out.append(_shell_folder(csidl) / LNK_NAME)
    if start_menu:
        csidl = 0x0002 if per_user else 0x0017        # PROGRAMS / COMMON_PROGRAMS
        out.append(_shell_folder(csidl) / APP_DIRNAME / LNK_NAME)
    return out


# ────────────────────────── поиск установленной копии ──────────────────────────

def find_existing_install(env=None) -> dict | None:
    """Ищет установленную копию: реестр, затем стандартные папки.

    Возвращает {'dir', 'build', 'version'} или None. Реестр и стандартные
    папки — только на Windows; в тестах проверяется ветка с явным списком
    папок через _probe_install_dir.
    """
    if IS_WINDOWS:
        try:
            import winreg
            for root, flags in ((winreg.HKEY_CURRENT_USER, 0),
                                (winreg.HKEY_LOCAL_MACHINE,
                                 winreg.KEY_READ | winreg.KEY_WOW64_32KEY),
                                (winreg.HKEY_LOCAL_MACHINE,
                                 winreg.KEY_READ | winreg.KEY_WOW64_64KEY)):
                try:
                    with winreg.OpenKey(root, UNINSTALL_REG, 0,
                                        winreg.KEY_READ | flags) as k:
                        loc, _ = winreg.QueryValueEx(k, "InstallLocation")
                except OSError:
                    continue
                info = _probe_install_dir(Path(loc))
                if info:
                    return info
        except Exception:
            pass
    for d in known_install_dirs(env):
        info = _probe_install_dir(d)
        if info:
            return info
    return None


def _probe_install_dir(d: Path) -> dict | None:
    d = Path(d)
    if not (d / EXE_NAME).is_file():
        return None
    build, version = 0, ""
    try:
        import app_update
        build = app_update.build_of_package(d) or 0
        version = app_update.version_of_package(d) or ""
    except Exception:
        pass
    return {"dir": str(d), "build": build, "version": version}


# ────────────────────────── установка ──────────────────────────

def _iter_files(src: Path):
    for p in sorted(src.rglob("*")):
        if p.is_file() and not p.name.lower().endswith((".lnk",)):
            yield p


def copy_program(src: Path, dest: Path, progress=None) -> int:
    """Копирует программу из src (временная распаковка) в dest.

    progress(done_bytes, total_bytes). Возвращает число скопированных
    файлов. Старые файлы прошлых сборок не трогаем — их подчистит сама
    программа при первом запуске (cleanup_stale_files по манифесту).
    """
    src, dest = Path(src), Path(dest)
    if not (src / EXE_NAME).is_file():
        raise FileNotFoundError("В источнике нет %s" % EXE_NAME)
    files = list(_iter_files(src))
    total = sum(f.stat().st_size for f in files)
    done = 0
    dest.mkdir(parents=True, exist_ok=True)
    for f in files:
        rel = f.relative_to(src)
        target = dest / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(f, target)
        done += f.stat().st_size
        if progress:
            try:
                progress(done, total)
            except Exception:
                pass
    return len(files)


def create_shortcut(lnk: Path, target: Path, args: str = "") -> None:
    """Создаёт .lnk через COM (IShellLinkW) — без pywin32."""
    if not IS_WINDOWS:
        raise NotImplementedError("ярлыки создаются только на Windows")
    import ctypes
    from ctypes import POINTER, byref, c_int, c_wchar_p
    from ctypes import oledll

    class GUID(ctypes.Structure):
        _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                    ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]

    CLSID_ShellLink = GUID(0x00021401, 0, 0,
                           (ctypes.c_ubyte * 8)(0xC0, 0, 0, 0, 0, 0, 0, 0x46))
    IID_IShellLinkW = GUID(0x000214F9, 0, 0,
                           (ctypes.c_ubyte * 8)(0xC0, 0, 0, 0, 0, 0, 0, 0x46))
    IID_IPersistFile = GUID(0x0000010B, 0, 0,
                            (ctypes.c_ubyte * 8)(0xC0, 0, 0, 0, 0, 0, 0, 0x46))

    # минимальный vtable IShellLinkW: QueryInterface..SetPath (слоты 0..19),
    # SetArguments — слот 11; IPersistFile: Save — слот 6
    slot = lambda obj, idx: ctypes.cast(
        ctypes.cast(obj, ctypes.c_void_p).value + idx * ctypes.sizeof(ctypes.c_void_p),
        POINTER(ctypes.c_void_p))

    punk = ctypes.c_void_p()
    oledll.ole32.CoInitializeEx(None, 0x2)  # COINIT_APARTMENTTHREADED
    try:
        oledll.ole32.CoCreateInstance(
            byref(CLSID_ShellLink), None, 1, byref(IID_IShellLinkW), byref(punk))
        # vtable-слоты после IUnknown (QI=0, AddRef=1, Release=2):
        # SetArguments — 11-й, SetPath — 18-й (последний метод IShellLinkW)
        fptr = ctypes.WINFUNCTYPE(c_int, ctypes.c_void_p, c_wchar_p)
        if args:
            fptr(slot(punk, 11)[0])(punk, str(args))
        fptr(slot(punk, 18)[0])(punk, str(target))
        ppf = ctypes.c_void_p()
        # QueryInterface(this, riid, ppv) — три аргумента
        qiptr = ctypes.WINFUNCTYPE(c_int, ctypes.c_void_p, ctypes.c_void_p,
                                   ctypes.POINTER(ctypes.c_void_p))
        qiptr(slot(punk, 0)[0])(punk, byref(IID_IPersistFile), byref(ppf))
        # IPersistFile::Save(pszFile, fRemember) — 6-й слот
        sptr = ctypes.WINFUNCTYPE(c_int, ctypes.c_void_p, c_wchar_p, c_int)
        sptr(slot(ppf, 6)[0])(ppf, str(lnk), 1)
    finally:
        oledll.ole32.CoUninitialize()


def write_uninstall_registry(install_dir: Path, per_user: bool,
                             version: str, build: int) -> None:
    """Запись в «Установка и удаление программ»."""
    if not IS_WINDOWS:
        return
    import winreg
    root = winreg.HKEY_CURRENT_USER if per_user else winreg.HKEY_LOCAL_MACHINE
    exe = Path(install_dir) / EXE_NAME
    with winreg.CreateKeyEx(root, UNINSTALL_REG, 0, winreg.KEY_WRITE) as k:
        winreg.SetValueEx(k, "DisplayName", 0, winreg.REG_SZ,
                          "OVERTIMETAB %s" % version)
        winreg.SetValueEx(k, "DisplayVersion", 0, winreg.REG_SZ,
                          "%s · сборка %d" % (version, build) if build else version)
        winreg.SetValueEx(k, "InstallLocation", 0, winreg.REG_SZ, str(install_dir))
        winreg.SetValueEx(k, "DisplayIcon", 0, winreg.REG_SZ, str(exe))
        winreg.SetValueEx(k, "UninstallString", 0, winreg.REG_SZ,
                          '"%s" --uninstall' % exe)
        winreg.SetValueEx(k, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(k, "NoRepair", 0, winreg.REG_DWORD, 1)


def do_install(src: Path, dest: Path, *, per_user: bool = True,
               desktop: bool = True, start_menu: bool = True,
               version: str = "", build: int = 0,
               progress=None) -> dict:
    """Полная установка. Возвращает отчёт {'dir', 'files', 'shortcuts'}."""
    dest = Path(dest).resolve()
    dest.mkdir(parents=True, exist_ok=True)
    # если программа сейчас запущена из этой папки — exe будет занят
    probe = dest / EXE_NAME
    if probe.exists():
        try:
            probe.rename(probe.with_name(EXE_NAME + ".probe"))
            probe.with_name(EXE_NAME + ".probe").rename(probe)
        except OSError:
            raise PermissionError(
                "Файл %s занят — закройте работающую программу и повторите."
                % probe)
    n = copy_program(src, dest, progress=progress)
    shortcuts = []
    if IS_WINDOWS and (desktop or start_menu):
        exe = dest / EXE_NAME
        for lnk in shortcut_paths(desktop, start_menu, per_user):
            try:
                create_shortcut(lnk, exe)
                shortcuts.append(str(lnk))
            except Exception:
                pass  # ярлык — не повод валить установку
    write_uninstall_registry(dest, per_user, version, build)
    return {"dir": str(dest), "files": n, "shortcuts": shortcuts}


# ────────────────────────── удаление ──────────────────────────

def remove_uninstall_registry() -> None:
    if not IS_WINDOWS:
        return
    import winreg
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            winreg.DeleteKey(root, UNINSTALL_REG)
        except OSError:
            pass


def remove_shortcuts() -> None:
    """Убираем все наши ярлыки (обоих типов размещения)."""
    if not IS_WINDOWS:
        return
    for lnk in shortcut_paths(True, True, True) + shortcut_paths(True, True, False):
        try:
            lnk.unlink()
        except OSError:
            pass
    # пустую папку в меню «Пуск» тоже приберём
    try:
        for per_user in (True, False):
            csidl = 0x0002 if per_user else 0x0017
            d = _shell_folder(csidl) / APP_DIRNAME
            if d.is_dir() and not any(d.iterdir()):
                d.rmdir()
    except Exception:
        pass


def schedule_dir_removal(path: Path) -> None:
    """Удаляет папку, из которой мы сами запущены (exe занят, пока жив процесс).

    Стандартный трюк: cmd ждёт пару секунд (пинг) и сносит папку, когда
    процесс уже завершился. На не-Windows — сразу shutil.
    """
    path = Path(path)
    if not IS_WINDOWS:
        shutil.rmtree(path, ignore_errors=True)
        return
    subprocess.Popen(
        ["cmd", "/c", "ping", "-n", "4", "127.0.0.1", ">nul",
         "&", "rmdir", "/s", "/q", str(path)],
        creationflags=0x08000000,  # CREATE_NO_WINDOW
        close_fds=True)


def do_uninstall(install_dir: Path, *, remove_data: bool = False,
                 home: str = "") -> dict:
    """Удаление: ярлыки, запись реестра, папка программы, (по галке) данные."""
    install_dir = Path(install_dir)
    remove_shortcuts()
    remove_uninstall_registry()
    if remove_data:
        shutil.rmtree(data_root(home), ignore_errors=True)
    schedule_dir_removal(install_dir)
    return {"dir": str(install_dir), "data_removed": bool(remove_data)}
