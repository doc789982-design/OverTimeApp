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
import zlib
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
    """Версия этой установленной копии. Ищем рядом с exe и в _internal."""
    here = Path(__file__).resolve().parent
    root = install_root(app_dir)
    candidates = [
        root / VERSION_JSON,
        root / "_internal" / VERSION_JSON,
        Path(app_dir) / VERSION_JSON,
        Path(app_dir) / "_internal" / VERSION_JSON,
        here / VERSION_JSON,
        here / THEME_REL,
        Path(app_dir) / THEME_REL,
    ]
    for p in candidates:
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
        try:
            with zipfile.ZipFile(path, "r") as zf:
                names = _zip_namelist_safe(zf)
            if _zip_has_exe(names):
                return path.resolve()
        except Exception:
            return None
        return None

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


def _norm_zip_name(name: str) -> str:
    return name.replace("\\", "/").lstrip("/")


def _zip_has_exe(names: list[str]) -> bool:
    for n in names:
        low = _norm_zip_name(n).lower()
        if low.endswith("overtimetab.exe") and not low.endswith("/"):
            return True
    return False


def _version_from_zip_bytes(raw: bytes, as_json: bool) -> str:
    text = raw.decode("utf-8", errors="replace")
    if as_json:
        try:
            return str(json.loads(text).get("version") or "").strip()
        except Exception:
            return ""
    return _read_version_from_text(text)


# GitHub zip (сборка --collect-all) часто без loose version.json:
# номер живёт внутри resources_rc (QML AppTheme). Имя файла zip не смотрим.
_JSON_VERSION_RE = re.compile(r'"version"\s*:\s*"([^"]+)"')
_ZIP_VERSION_CACHE: dict[str, tuple[int, int, str]] = {}
_VERSION_MEMBER_HINTS = (
    "version.json",
    "apptheme.qml",
    "changelog.md",
    "resources_rc.py",
    "resources_rc.pyc",
    "version_info.py",
)


def _looks_like_app_version(ver: str) -> bool:
    ver = (ver or "").strip()
    if not ver:
        return False
    return bool(re.match(r"\d+\.\d+", ver))


def _extract_version_from_plain(raw: bytes) -> str:
    if not raw:
        return ""
    text = raw.decode("utf-8", errors="replace")
    ver = _read_version_from_text(text)
    if _looks_like_app_version(ver):
        return ver
    try:
        data = json.loads(text)
        ver = str(data.get("version") or "").strip()
        if _looks_like_app_version(ver):
            return ver
    except Exception:
        pass
    m = _JSON_VERSION_RE.search(text)
    if m and _looks_like_app_version(m.group(1)):
        cand = m.group(1).strip()
        if _looks_like_app_version(cand) and "pyside" not in cand.lower():
            return cand
    m = re.search(r"(?m)^##\s+(\d+\.\d+\S*)", text)
    if m and _looks_like_app_version(m.group(1)):
        return m.group(1).strip()
    m = re.search(r"OVERTIMETAB_APP_VERSION\s*=\s*[\"']([^\"']+)[\"']", text)
    if m and _looks_like_app_version(m.group(1)):
        return m.group(1).strip()
    return ""


def _extract_version_from_bytes(raw: bytes) -> str:
    if not raw:
        return ""
    ver = _extract_version_from_plain(raw)
    if ver:
        return ver
    # pyside6-rcc пишет байты как \x7b\x22version\x22 — снимаем экранирование
    try:
        src = raw.decode("latin-1", errors="replace")
        if "\\x" in src:
            unescaped = re.sub(
                r"\\x([0-9a-fA-F]{2})",
                lambda m: chr(int(m.group(1), 16)),
                src,
            )
            ver = _extract_version_from_plain(unescaped.encode("latin-1", errors="replace"))
            if ver:
                return ver
    except Exception:
        pass
    # PYZ / qCompress: zlib-потоки внутри файла
    tries = 0
    for m in re.finditer(b"\\x78[\\x01\\x5e\\x9c\\xda]", raw):
        i = m.start()
        if i > 8 * 1024 * 1024:
            break
        try:
            dec = zlib.decompress(raw[i : i + 2 * 1024 * 1024])
        except Exception:
            continue
        ver = _extract_version_from_plain(dec)
        if ver:
            return ver
        tries += 1
        if tries >= 24:
            break
    return ""


def _zip_member_priority(name: str) -> tuple:
    low = name.lower()
    if low.endswith("version.json") and "/data/" not in ("/" + low):
        return (0, low.count("/"), len(low))
    if low.endswith("apptheme.qml"):
        return (1, low.count("/"), len(low))
    if "resources_rc" in Path(low).name:
        return (2, low.count("/"), len(low))
    if low.endswith("changelog.md"):
        return (3, low.count("/"), len(low))
    return (9, low.count("/"), len(low))


def _iter_version_members(names: list[str]) -> list[str]:
    hits = []
    for n in names:
        low = n.lower()
        if "/data/" in ("/" + low):
            continue
        base = Path(low).name
        if any(base.endswith(h) or h in base for h in _VERSION_MEMBER_HINTS):
            hits.append(n)
            continue
        if "resources_rc" in base or base.endswith(".pyz") or "pyz-" in base:
            hits.append(n)
            continue
        if base == "overtimetab.exe":
            hits.append(n)
    hits.sort(key=_zip_member_priority)
    return hits


def _version_from_zipfile(zf: zipfile.ZipFile, names: list[str]) -> str:
    for n in _iter_version_members(names):
        try:
            info = zf.getinfo(n)
        except KeyError:
            # namelist was normalised; find original
            info = None
            for cand in zf.infolist():
                if _norm_zip_name(cand.filename) == n:
                    info = cand
                    break
            if info is None:
                continue
        base = Path(n).name.lower()
        limit = 40 * 1024 * 1024 if ("pyz" in base or "resources_rc" in base or base.endswith(".exe")) else 12 * 1024 * 1024
        if info.file_size > limit:
            continue
        try:
            with zf.open(info) as fh:
                raw = fh.read()
            ver = _extract_version_from_bytes(raw)
            if ver:
                return ver
        except Exception:
            continue
    return ""


def version_of_package(package: Path) -> str:
    """Номер версии изнутри посылки. Имя файла zip не смотрим."""
    package = Path(package)
    if package.is_file() and package.suffix.lower() == ".zip":
        try:
            st = package.stat()
            key = str(package.resolve())
            cached = _ZIP_VERSION_CACHE.get(key)
            if cached and cached[0] == st.st_mtime_ns and cached[1] == st.st_size:
                return cached[2]
        except Exception:
            key = ""
            st = None
        ver = ""
        try:
            with zipfile.ZipFile(package, "r") as zf:
                names = [_norm_zip_name(n) for n in _zip_namelist_safe(zf)]
                ver = _version_from_zipfile(zf, names)
        except Exception:
            ver = ""
        if key and st is not None:
            _ZIP_VERSION_CACHE[key] = (st.st_mtime_ns, st.st_size, ver)
        return ver

    if package.is_dir():
        for cand in (
            package / VERSION_JSON,
            package / "_internal" / VERSION_JSON,
            package / THEME_REL,
            package / "_internal" / THEME_REL,
        ):
            if not cand.exists():
                continue
            ver = read_version_json(cand) if cand.suffix.lower() == ".json" else read_version_from_theme_file(cand)
            if ver:
                return ver
        # Сборка без version.json: номер внутри resources_rc (как в zip с GitHub).
        extra: list[Path] = []
        for base in (package, package / "_internal"):
            try:
                if not base.is_dir():
                    continue
                for child in base.iterdir():
                    if not child.is_file():
                        continue
                    low = child.name.lower()
                    if low in ("changelog.md", "version.json") or "resources_rc" in low or "pyz" in low:
                        extra.append(child)
            except Exception:
                continue
        for cand in extra:
            try:
                limit = 40 * 1024 * 1024 if "pyz" in cand.name.lower() or "resources_rc" in cand.name.lower() else 12 * 1024 * 1024
                if cand.stat().st_size > limit:
                    continue
                ver = _extract_version_from_bytes(cand.read_bytes())
            except Exception:
                continue
            if ver:
                return ver
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


def cleanup_obsolete_zips(app_dir: Path, current_version: str) -> list[str]:
    """
    После установки: рядом с exe удаляем zip самой программы,
    если внутри уже не новая версия. Чужие архивы не трогаем.
    Флешку не чистим — только папка с программой.
    """
    removed: list[str] = []
    if not current_version:
        return removed
    root = install_root(app_dir)
    folders = {root}
    try:
        folders.add(Path(app_dir).resolve())
    except Exception:
        pass
    for folder in folders:
        try:
            if not folder.is_dir():
                continue
        except Exception:
            continue
        try:
            children = list(folder.iterdir())
        except Exception:
            continue
        for child in children:
            try:
                if not child.is_file() or child.suffix.lower() != ".zip":
                    continue
                if find_package_root(child) is None:
                    continue
                ver = version_of_package(child)
                if not ver:
                    continue
                if is_newer(ver, current_version):
                    continue
                child.unlink()
                removed.append(str(child))
            except Exception:
                continue
    return removed


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


def _is_drive_root(path: Path) -> bool:
    try:
        path = Path(path).resolve()
        return path.parent == path
    except Exception:
        return False


def _user_drop_dirs() -> list[Path]:
    """Рабочий стол и Загрузки: zip часто кладут рядом с ярлыком, а не с настоящим exe."""
    dirs: list[Path] = []
    try:
        home = Path.home()
    except Exception:
        return dirs
    names = (
        "Desktop",
        "Downloads",
        "OneDrive/Desktop",
        "OneDrive/Downloads",
        "Рабочий стол",
        "Загрузки",
    )
    for name in names:
        p = home / name
        try:
            if p.is_dir():
                dirs.append(p)
        except Exception:
            continue
    return dirs


def scan_update_sources(app_dir: Path) -> list[Path]:
    """Кандидаты рядом с exe и на флешках. Имя файла не важно — смотрим содержимое."""
    found: list[Path] = []
    app_dir = Path(app_dir)

    staged = staged_info(app_dir)
    if staged:
        found.append(Path(staged["root"]))

    root = install_root(app_dir)
    skip_roots = {app_dir.resolve(), root.resolve(), pending_dir(app_dir).resolve()}
    search_dirs = [root, app_dir]
    if not _is_drive_root(root.parent):
        search_dirs.append(root.parent)
    if not _is_drive_root(app_dir.parent):
        search_dirs.append(app_dir.parent)
    if getattr(sys, "frozen", False):
        try:
            search_dirs.append(Path(sys.executable).resolve().parent)
        except Exception:
            pass
    search_dirs.extend(_user_drop_dirs())
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
                    if child.is_file() and child.suffix.lower() == ".zip":
                        pkg = find_package_root(child)
                        if not pkg:
                            continue
                        rp = pkg.resolve()
                        if rp not in seen:
                            found.append(rp)
                            seen.add(rp)
                    elif child.is_dir():
                        if child.name.lower() in SKIP_NAMES:
                            continue
                        pkg = find_package_root(child)
                        if not pkg:
                            continue
                        rp = pkg.resolve()
                        if rp not in seen and rp not in skip_roots:
                            found.append(rp)
                            seen.add(rp)
                    elif child.is_file() and child.name.lower() == EXE_NAME.lower():
                        rp = child.parent.resolve()
                        if rp not in seen and rp not in skip_roots:
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
# Переодевание: скрытый VBS в TEMP.
# bat + tasklist|findstr открывал чёрное окно и зацикливался:
# у канала cmd errorlevel берётся от tasklist (всегда 0), а не от findstr.
# ---------------------------------------------------------------------------

# Одинарные тройные кавычки: в VBS " и "" обычные, а """ внутри
# r"""...""" обрывает строку Python и получается NameError: src.
_VBS_TEMPLATE = r'''Option Explicit
Dim src, dst, pid, exe, sh, fso, logFile, t0, rc, q
q = Chr(34)
src = WScript.Arguments(0)
dst = WScript.Arguments(1)
pid = WScript.Arguments(2)
exe = dst & "\OVERTIMETAB.exe"

Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
logFile = sh.ExpandEnvironmentStrings("%TEMP%") & "\overtimetab_update.log"
WriteLog "start src=" & src & " dst=" & dst & " pid=" & pid

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

rc = sh.Run("robocopy " & q & src & q & " " & q & dst & q & " /E /XD data pending_update /XF *.sqlite *.sqlite-wal *.sqlite-shm /NFL /NDL /NJH /NJS /NC /NS /NP /R:3 /W:1", 0, True)
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

On Error Resume Next
If fso.FolderExists(src) Then fso.DeleteFolder src, True
fso.DeleteFile WScript.ScriptFullName, True
On Error GoTo 0
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


def launch_file_swap(source: Path, dest: Path, pid: int) -> Path:
    """Пишет скрытый .vbs в TEMP и запускает его без консоли."""
    source = Path(source).resolve()
    dest = Path(dest).resolve()
    if not _exe_in(source):
        raise UpdateError("Подготовленное обновление повреждено")
    if os.name != "nt":
        raise UpdateError("Переодевание файлов рассчитано на Windows")

    vbs_path = Path(tempfile.gettempdir()) / f"ot_update_{os.getpid()}.vbs"
    vbs_path.write_text(_VBS_TEMPLATE, encoding="ascii", errors="replace")

    windir = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    wscript = windir / "System32" / "wscript.exe"
    if not wscript.exists():
        wscript = Path("wscript.exe")

    creation = 0x08000000  # CREATE_NO_WINDOW, без DETACHED_PROCESS
    startup = None
    if hasattr(subprocess, "STARTUPINFO"):
        startup = subprocess.STARTUPINFO()
        startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startup.wShowWindow = 0

    subprocess.Popen(
        [str(wscript), "//B", "//Nologo", str(vbs_path), str(source), str(dest), str(pid)],
        cwd=str(Path(tempfile.gettempdir())),
        close_fds=True,
        creationflags=creation,
        startupinfo=startup,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return vbs_path


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


# ---------------------------------------------------------------------------
# Журнал изменений (показываем после установки)
# ---------------------------------------------------------------------------

_CHANGELOG_HEADER_RE = re.compile(r"^##\s+(\S+)(?:\s+[—–-]\s+(.+))?\s*$")
_CHANGELOG_SECTION_RE = re.compile(r"^###\s+(.+?)\s*$")
_CHANGELOG_SECTION_MAP = {
    "добавили": "added",
    "поменяли": "changed",
    "починили": "fixed",
    "удалили": "removed",
    "сборка": "build",
}


def parse_changelog(text: str) -> list[dict]:
    """Разбирает CHANGELOG.md. Сверху файла — самое новое."""
    blocks: list[dict] = []
    current = None
    section = None
    for raw_line in (text or "").splitlines():
        line = raw_line.rstrip()
        m = _CHANGELOG_HEADER_RE.match(line)
        if m:
            if current:
                blocks.append(current)
            current = {
                "version": m.group(1).strip(),
                "date": (m.group(2) or "").strip(),
                "added": [],
                "changed": [],
                "fixed": [],
                "removed": [],
                "build": [],
            }
            section = None
            continue
        if not current:
            continue
        sm = _CHANGELOG_SECTION_RE.match(line)
        if sm:
            section = _CHANGELOG_SECTION_MAP.get(sm.group(1).strip().lower())
            continue
        if line.startswith("---") or not line.strip():
            continue
        if not section:
            continue
        if line.startswith("- "):
            current[section].append(line[2:].strip())
        elif line.startswith("  "):
            if current[section]:
                current[section][-1] += " " + line.strip()
        elif not line.startswith("#"):
            if current[section]:
                current[section][-1] += " " + line.strip()
            else:
                current[section].append(line.strip())
    if current:
        blocks.append(current)
    return blocks


def changelog_since(text: str, from_version: str, to_version: str) -> list[dict]:
    """
    Записи новее from_version, не новее to_version.
    Если from_version пустой — только текущая (чтобы не вывалить всю историю).
    """
    out: list[dict] = []
    for b in parse_changelog(text):
        ver = (b.get("version") or "").strip()
        if not ver:
            continue
        if to_version and ver != to_version and not is_newer(to_version, ver):
            continue
        if from_version:
            if not is_newer(ver, from_version):
                continue
        elif to_version and ver != to_version:
            continue
        if not (b["added"] or b["changed"] or b["fixed"] or b["removed"]):
            continue
        out.append(b)
    return out


def _bullets(items: list[str]) -> str:
    return "\n".join("• " + x for x in items if x)


def changelog_for_qml(blocks: list[dict]) -> list[dict]:
    """Плоские словари для QML Repeater."""
    out = []
    for b in blocks:
        added = list(b.get("added") or [])
        changed = list(b.get("changed") or [])
        fixed = list(b.get("fixed") or [])
        removed = list(b.get("removed") or [])
        out.append({
            "version": b.get("version") or "",
            "date": b.get("date") or "",
            "hasAdded": bool(added),
            "hasChanged": bool(changed),
            "hasFixed": bool(fixed),
            "hasRemoved": bool(removed),
            "addedText": _bullets(added),
            "changedText": _bullets(changed),
            "fixedText": _bullets(fixed),
            "removedText": _bullets(removed),
        })
    return out


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
