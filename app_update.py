# -*- coding: utf-8 -*-
"""
Офлайн-обновление OVERTIMETAB (как кнопка в Telegram, без интернета).

Программа и данные живут рядом, но это разные вещи:
  data/              — базы, хоткеи, тема. При обновлении НЕ трогаем.
  всё остальное      — «коробка». Её можно заменить.

Новая версия приходит zip-архивом или папкой (флешка, общая папка).
Готовим её в pending_update/, показываем кнопку, по клику
маленький .bat в %TEMP% ждёт закрытия программы, копирует файлы
и снова запускает OVERTIMETAB.exe.

Этот модуль без Qt и без logic.py — только файлы и версии.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Optional


EXE_NAME = "OVERTIMETAB.exe"
PENDING_DIRNAME = "pending_update"
DATA_DIRNAME = "data"
META_NAME = "UPDATE_META.json"
VERSION_JSON = "version.json"
THEME_REL = Path("components") / "AppTheme.qml"

VERSION_RE = re.compile(r'appVersion:\s*"([^"]+)"')

# Эти папки/файлы никогда не переезжают из новой коробки поверх старой.
SKIP_NAMES = {
    DATA_DIRNAME,
    PENDING_DIRNAME,
    "__pycache__",
    ".git",
    "data",
}

# Имена zip/папок, которые считаем «посылкой с обновлением»
ZIP_PREFIXES = ("overtimetab", "overtimeapp")


class UpdateError(Exception):
    pass


# ---------------------------------------------------------------------------
# Версии
# ---------------------------------------------------------------------------

def parse_version(raw: str) -> tuple:
    """
    '2.0.0-ALPHA.19' -> ((2, 0, 0), 'ALPHA', 19)
    '2.0.1'          -> ((2, 0, 1), '', 0)
    Пустой пререлиз (релиз) старше любого ALPHA/BETA.
    """
    s = (raw or "").strip()
    if not s:
        return ((0, 0, 0), "", 0)
    core, _, tail = s.partition("-")
    nums = []
    for part in core.split("."):
        try:
            nums.append(int(part))
        except ValueError:
            nums.append(0)
    while len(nums) < 3:
        nums.append(0)
    pre_name, pre_num = "", 0
    if tail:
        name, _, num = tail.partition(".")
        pre_name = name.upper()
        try:
            pre_num = int(num) if num else 0
        except ValueError:
            pre_num = 0
    return (tuple(nums[:3]), pre_name, pre_num)


def is_newer(candidate: str, current: str) -> bool:
    """candidate новее current?"""
    if not candidate:
        return False
    if not current:
        return True
    c, v = parse_version(candidate), parse_version(current)
    if c[0] != v[0]:
        return c[0] > v[0]
    # одинаковая тройка: релиз (без хвоста) новее любой альфы
    if not c[1] and v[1]:
        return True
    if c[1] and not v[1]:
        return False
    if c[1] != v[1]:
        return c[1] > v[1]
    return c[2] > v[2]


def _read_version_from_text(text: str) -> str:
    m = VERSION_RE.search(text or "")
    return m.group(1) if m else ""


def read_version_from_theme_file(path: Path) -> str:
    try:
        return _read_version_from_text(path.read_text(encoding="utf-8"))
    except Exception:
        return ""


def read_version_json(path: Path) -> str:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return str(data.get("version") or "").strip()
    except Exception:
        return ""


def write_version_json(path: Path, version: str) -> None:
    path.write_text(
        json.dumps({"name": "OVERTIMETAB", "version": version}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def current_app_version(app_dir: Path) -> str:
    """Версия этой установленной копии. Ищем рядом с exe и в исходниках."""
    here = Path(__file__).resolve().parent
    root = install_root(app_dir)
    for p in (
        root / VERSION_JSON,
        app_dir / VERSION_JSON,
        here / VERSION_JSON,
        here / THEME_REL,
        app_dir / THEME_REL,
    ):
        if not p.exists():
            continue
        ver = read_version_json(p) if p.suffix.lower() == ".json" else read_version_from_theme_file(p)
        if ver:
            return ver
    return ""


# ---------------------------------------------------------------------------
# Распознавание пакета
# ---------------------------------------------------------------------------

def _exe_in(folder: Path) -> bool:
    return (folder / EXE_NAME).exists() or (folder / EXE_NAME.lower()).exists()


def find_package_root(path: Path) -> Optional[Path]:
    """
    Папка с OVERTIMETAB.exe или сам .zip.
    Понимает и «содержимое dist/OVERTIMETAB», и обёртку OVERTIMETAB/OVERTIMETAB.exe.
    """
    if not path:
        return None
    try:
        path = Path(path).expanduser()
        if not path.exists():
            return None
    except Exception:
        return None

    if path.is_file() and path.suffix.lower() == ".zip":
        return path.resolve()

    if path.is_file() and path.name.lower() == EXE_NAME.lower():
        return path.parent.resolve()

    if path.is_dir():
        path = path.resolve()
        if _exe_in(path):
            return path
        inner = path / "OVERTIMETAB"
        if inner.is_dir() and _exe_in(inner):
            return inner
    return None


def _zip_namelist_safe(zf: zipfile.ZipFile) -> list[str]:
    names = []
    for info in zf.infolist():
        name = info.filename.replace("\\", "/")
        if name.startswith("/") or name.startswith("\\") or ".." in Path(name).parts:
            raise UpdateError("В архиве подозрительные пути — обновление отклонено")
        names.append(name)
    return names


def version_of_package(package: Path) -> str:
    """Версия посылки, не распаковывая целиком."""
    package = Path(package)
    if package.is_file() and package.suffix.lower() == ".zip":
        with zipfile.ZipFile(package, "r") as zf:
            names = _zip_namelist_safe(zf)
            for key in (VERSION_JSON, "OVERTIMETAB/" + VERSION_JSON):
                if key in names:
                    with zf.open(key) as fh:
                        data = json.loads(fh.read().decode("utf-8"))
                        return str(data.get("version") or "").strip()
            for name in names:
                if name.endswith("components/AppTheme.qml") or name.endswith("AppTheme.qml"):
                    with zf.open(name) as fh:
                        ver = _read_version_from_text(fh.read().decode("utf-8", errors="replace"))
                        if ver:
                            return ver
        return ""

    if package.is_dir():
        v = read_version_json(package / VERSION_JSON)
        if v:
            return v
        return read_version_from_theme_file(package / THEME_REL)
    return ""


def describe_package(package: Path) -> dict:
    root = find_package_root(package)
    if not root:
        raise UpdateError("Это не папка и не архив OVERTIMETAB")
    ver = version_of_package(root)
    return {"root": str(root), "version": ver, "is_zip": root.is_file()}


# ---------------------------------------------------------------------------
# Подготовка (staging)
# ---------------------------------------------------------------------------

def install_root(app_dir: Path) -> Path:
    """Папка с OVERTIMETAB.exe. В PyInstaller 6 это родитель _internal."""
    app_dir = Path(app_dir).resolve()
    if (app_dir / EXE_NAME).exists():
        return app_dir
    parent = app_dir.parent
    if (parent / EXE_NAME).exists():
        return parent
    if getattr(sys, "frozen", False):
        try:
            return Path(sys.executable).resolve().parent
        except Exception:
            pass
    return app_dir


def pending_dir(app_dir: Path) -> Path:
    return install_root(app_dir) / PENDING_DIRNAME


def _clear_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)
    if path.exists():
        # на Windows иногда держит антивирус — пробуем ещё раз
        time.sleep(0.3)
        shutil.rmtree(path, ignore_errors=False)


def _copy_tree(src: Path, dst: Path) -> None:
    def ignore(directory, names):
        skipped = []
        for n in names:
            if n in SKIP_NAMES or n.endswith(".sqlite") or n.endswith(".sqlite-wal") or n.endswith(".sqlite-shm"):
                skipped.append(n)
        return set(skipped)

    shutil.copytree(src, dst, dirs_exist_ok=True, ignore=ignore)


def _extract_zip(zip_path: Path, dest: Path) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as zf:
        names = _zip_namelist_safe(zf)
        # Если в корне архива папка OVERTIMETAB/ — снимаем этот префикс
        prefix = ""
        exe_hits = [n for n in names if n.replace("\\", "/").lower().endswith("overtimetab.exe") and not n.endswith("/")]
        if not exe_hits:
            raise UpdateError("В архиве нет OVERTIMETAB.exe")
        first = exe_hits[0].replace("\\", "/")
        parts = first.split("/")
        if len(parts) >= 2 and parts[0].lower() == "overtimetab":
            prefix = parts[0] + "/"

        for info in zf.infolist():
            name = info.filename.replace("\\", "/")
            rel = name[len(prefix):] if prefix and name.startswith(prefix) else name
            if not rel or rel.endswith("/"):
                continue
            parts = [p for p in rel.split("/") if p]
            if any(p in SKIP_NAMES or p.lower() == DATA_DIRNAME for p in parts):
                continue
            target = dest / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)


def stage_package(package: Path, app_dir: Path) -> str:
    """
    Кладёт новую коробку в app_dir/pending_update.
    Возвращает версию (может быть пустой, если в пакете нет version.json).
    """
    root = find_package_root(package)
    if not root:
        raise UpdateError("Не нашли OVERTIMETAB.exe в выбранном месте")

    app_dir = Path(app_dir).resolve()
    dest_root = install_root(app_dir)
    try:
        if root.resolve() in (app_dir, dest_root):
            raise UpdateError("Выбрана папка, в которой программа уже запущена")
    except UpdateError:
        raise
    except Exception:
        pass

    ver = version_of_package(root)
    dest = pending_dir(app_dir)
    tmp = dest.with_name(dest.name + ".__tmp")
    _clear_dir(tmp)
    tmp.mkdir(parents=True, exist_ok=True)

    try:
        if root.is_file() and root.suffix.lower() == ".zip":
            _extract_zip(root, tmp)
        else:
            _copy_tree(root, tmp)

        if not _exe_in(tmp):
            raise UpdateError("После распаковки не оказалось OVERTIMETAB.exe")

        write_version_json(tmp / VERSION_JSON, ver) if ver else None
        meta = {"version": ver, "source": str(root)}
        (tmp / META_NAME).write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

        _clear_dir(dest)
        tmp.rename(dest)
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise

    return ver


def staged_info(app_dir: Path) -> Optional[dict]:
    dest = pending_dir(app_dir)
    if not dest.is_dir() or not _exe_in(dest):
        return None
    ver = ""
    meta = dest / META_NAME
    if meta.exists():
        try:
            ver = str(json.loads(meta.read_text(encoding="utf-8")).get("version") or "")
        except Exception:
            ver = ""
    if not ver:
        ver = version_of_package(dest)
    return {"root": str(dest), "version": ver}


def cleanup_pending(app_dir: Path) -> None:
    dest = pending_dir(app_dir)
    if dest.exists():
        shutil.rmtree(dest, ignore_errors=True)


# ---------------------------------------------------------------------------
# Поиск посылки (флешка, рядом с программой)
# ---------------------------------------------------------------------------

def _iter_removable_roots() -> list[Path]:
    roots: list[Path] = []
    if os.name != "nt":
        return roots
    try:
        import ctypes
        kernel = ctypes.windll.kernel32  # type: ignore[attr-defined]
        buf = ctypes.create_unicode_buffer(512)
        n = kernel.GetLogicalDriveStringsW(511, buf)
        raw = buf[:n]
        drives = [d for d in raw.split("\x00") if d]
        DRIVE_REMOVABLE = 2
        DRIVE_FIXED = 3
        for d in drives:
            dtype = kernel.GetDriveTypeW(d)
            # Съемные точно, плюс фиксированные кроме системного диска —
            # флешка иногда монтируется как Local Disk.
            if dtype == DRIVE_REMOVABLE:
                roots.append(Path(d))
            elif dtype == DRIVE_FIXED:
                sys_drive = os.environ.get("SystemDrive", "C:") + "\\"
                if d.upper().rstrip("\\") != sys_drive.upper().rstrip("\\"):
                    roots.append(Path(d))
    except Exception:
        pass
    return roots


def _looks_like_update_zip(path: Path) -> bool:
    name = path.name.lower()
    if path.suffix.lower() != ".zip":
        return False
    return any(name.startswith(p) for p in ZIP_PREFIXES) or "overtime" in name


def scan_update_sources(app_dir: Path) -> list[Path]:
    """Кандидаты: уже подготовленная папка, zip рядом, флешки."""
    found: list[Path] = []
    app_dir = Path(app_dir)

    staged = staged_info(app_dir)
    if staged:
        found.append(Path(staged["root"]))

    root = install_root(app_dir)
    search_dirs = [root, root.parent, app_dir, app_dir.parent]
    search_dirs.extend(_iter_removable_roots())

    seen = {p.resolve() for p in found}
    for folder in search_dirs:
        try:
            if not folder.exists() or not folder.is_dir():
                continue
        except Exception:
            continue
        try:
            for child in folder.iterdir():
                try:
                    if child.is_file() and _looks_like_update_zip(child):
                        rp = child.resolve()
                        if rp not in seen:
                            found.append(rp)
                            seen.add(rp)
                    elif child.is_dir() and child.name.lower() in ("overtimetab", "overtimeapp"):
                        pkg = find_package_root(child)
                        if pkg:
                            rp = pkg.resolve()
                            if rp not in seen and rp not in (app_dir.resolve(), root):
                                found.append(rp)
                                seen.add(rp)
                    elif child.is_file() and child.name.lower() == EXE_NAME.lower():
                        rp = child.parent.resolve()
                        if rp not in seen and rp not in (app_dir.resolve(), root):
                            found.append(rp)
                            seen.add(rp)
                except Exception:
                    continue
        except Exception:
            continue
    return found


def pick_best_update(app_dir: Path, current_version: str) -> Optional[dict]:
    """Самая новая посылка, которая новее текущей. Без версии — пропускаем при автопоиске."""
    best = None
    best_ver = current_version
    for src in scan_update_sources(app_dir):
        try:
            if src.resolve() == pending_dir(app_dir).resolve():
                info = staged_info(app_dir)
                if not info:
                    continue
                ver = info["version"]
                if ver and is_newer(ver, current_version):
                    return {"root": info["root"], "version": ver, "already_staged": True}
                continue
            ver = version_of_package(src)
            if not ver:
                continue
            if is_newer(ver, best_ver if best else current_version):
                best = {"root": str(src), "version": ver, "already_staged": False}
                best_ver = ver
        except Exception:
            continue
    return best


# ---------------------------------------------------------------------------
# Переодевание: bat в TEMP, потому что Windows не даст перезаписать свой exe
# ---------------------------------------------------------------------------

_BAT_TEMPLATE = r"""@echo off
setlocal EnableExtensions
set "SRC=%~1"
set "DST=%~2"
set "PID=%~3"
set "EXE=%DST%\OVERTIMETAB.exe"
set "LOG=%TEMP%\overtimetab_update.log"

echo start %DATE% %TIME% > "%LOG%"
echo src=%SRC%>> "%LOG%"
echo dst=%DST%>> "%LOG%"
echo pid=%PID%>> "%LOG%"

:wait
tasklist /FI "PID eq %PID%" 2>nul | findstr /I /C:"%PID%" >nul
if not errorlevel 1 (
  ping -n 2 127.0.0.1 >nul
  goto wait
)

ping -n 2 127.0.0.1 >nul

if not exist "%SRC%\OVERTIMETAB.exe" (
  echo missing source exe>> "%LOG%"
  exit /b 1
)

robocopy "%SRC%" "%DST%" /E /XD data pending_update /XF *.sqlite *.sqlite-wal *.sqlite-shm /NFL /NDL /NJH /NJS /nc /ns /np /R:3 /W:1
set RC=%ERRORLEVEL%
echo robocopy=%RC%>> "%LOG%"

if %RC% GEQ 8 (
  echo robocopy failed>> "%LOG%"
  if exist "%EXE%" start "" "%EXE%"
  exit /b 1
)

if exist "%EXE%" (
  start "" "%EXE%"
  echo restarted>> "%LOG%"
)

rmdir /s /q "%SRC%" 2>nul
del "%~f0" >nul 2>&1
"""


def launch_file_swap(source: Path, dest: Path, pid: int) -> Path:
    """Пишет .bat в TEMP и запускает его отвязанным процессом."""
    source = Path(source).resolve()
    dest = Path(dest).resolve()
    if not _exe_in(source):
        raise UpdateError("Подготовленное обновление повреждено")

    bat_path = Path(tempfile.gettempdir()) / f"ot_update_{os.getpid()}.bat"
    bat_path.write_text(_BAT_TEMPLATE, encoding="ascii", errors="replace")

    creation = 0
    if os.name == "nt":
        creation = 0x00000008 | 0x08000000  # DETACHED_PROCESS | CREATE_NO_WINDOW

    subprocess.Popen(
        ["cmd.exe", "/c", str(bat_path), str(source), str(dest), str(pid)],
        cwd=str(Path(tempfile.gettempdir())),
        close_fds=True,
        creationflags=creation,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return bat_path


def apply_update_inplace(source: Path, dest: Path, wait_pid: int | None = None) -> None:
    """Запасной путь: нас запустили из pending_update с --apply-update."""
    if wait_pid:
        deadline = time.time() + 60
        while time.time() < deadline:
            if not _pid_alive(wait_pid):
                break
            time.sleep(0.4)
        time.sleep(0.6)

    source = Path(source).resolve()
    dest = Path(dest).resolve()
    if source == dest:
        raise UpdateError("Источник и назначение совпадают")

    for item in source.iterdir():
        if item.name in SKIP_NAMES:
            continue
        target = dest / item.name
        if item.is_dir():
            if target.exists():
                shutil.rmtree(target, ignore_errors=True)
            shutil.copytree(item, target, ignore=lambda d, names: {n for n in names if n in SKIP_NAMES})
        else:
            shutil.copy2(item, target)

    exe = dest / EXE_NAME
    if exe.exists():
        creation = 0x00000008 if os.name == "nt" else 0
        subprocess.Popen([str(exe)], cwd=str(dest), creationflags=creation, close_fds=True)


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if os.name == "nt":
        try:
            out = subprocess.check_output(
                ["tasklist", "/FI", f"PID eq {pid}"],
                stderr=subprocess.DEVNULL,
                creationflags=0x08000000 if os.name == "nt" else 0,
            )
            return str(pid).encode() in out
        except Exception:
            return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def parse_apply_argv(argv: list[str]) -> Optional[dict]:
    if "--apply-update" not in argv:
        return None
    out = {"source": "", "dest": "", "wait_pid": 0}
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--from" and i + 1 < len(argv):
            out["source"] = argv[i + 1]
            i += 2
            continue
        if a == "--to" and i + 1 < len(argv):
            out["dest"] = argv[i + 1]
            i += 2
            continue
        if a == "--wait-pid" and i + 1 < len(argv):
            try:
                out["wait_pid"] = int(argv[i + 1])
            except ValueError:
                out["wait_pid"] = 0
            i += 2
            continue
        i += 1
    return out
