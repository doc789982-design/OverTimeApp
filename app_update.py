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

import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import zlib
from pathlib import Path
from typing import Optional


EXE_NAME = "OVERTIMETAB.exe"
PENDING_DIRNAME = "pending_update"
DATA_DIRNAME = "data"
META_NAME = "UPDATE_META.json"
ROLLBACK_DIRNAME = "rollback"
ROLLBACK_META_NAME = "ROLLBACK_META.json"
VERSION_JSON = "version.json"
THEME_REL = Path("components") / "AppTheme.qml"

# Постоянный адрес хранилища обновлений.
# Ссылка на GitHub-релиз OverTimeTab: программа проверяет обновления сама,
# ничего вставлять в настройках не нужно. Если пользователь всё же ввёл свой
# адрес в поле настроек — он имеет приоритет (хранится в config.json и
# переживает перезапуск и обновление). resolve_update_url() отдаёт введённый
# адрес, а если он пуст — этот постоянный.
DEFAULT_UPDATE_URL = "https://github.com/doc789982-design/OverTimeApp/releases/latest"

# Второй вшитый адрес хранилища: если GitHub недоступен (например, в закрытой
# сети), запрос пробивается сюда. Формат — как у любого веб-хранилища:
# рядом с архивом лежит version.json (его готовит tools/make_web_release.py).
# Порядок проверки обновлений: адрес из настроек (если указан) → GitHub → сюда.
# Адрес ровно в том виде, в каком его дал владелец программы (http).
# Не «улучшать» до https: на защищённых рабочих местах https туда может
# не пройти (проверка сертификатов), а браузеры качают по http.
FALLBACK_UPDATE_URL = "http://post.mvd.ru/~mgrigorev46@mvd.ru/"


def resolve_update_url(stored_url: str = "") -> str:
    """Адрес хранилища: введённый в настройках, а если пусто — постоянный."""
    return (stored_url or "").strip() or DEFAULT_UPDATE_URL

VERSION_RE = re.compile(r'appVersion:\s*"([^"]+)"')

# Эти папки/файлы никогда не переезжают из новой коробки поверх старой.
SKIP_NAMES = {
    DATA_DIRNAME,
    PENDING_DIRNAME,
    "__pycache__",
    ".git",
    "data",
}

# Манифест файлов сборки (готовится tools/overtimetab.spec): программа при
# старте сверяет с ним свою папку и убирает файлы прошлых версий —
# установка «поверх» старой сборки не оставляет мусора.
MANIFEST_NAME = "app_files.txt"
CLEANUP_MARKER = "cleanup_build.txt"

# Что никогда не убираем даже вне манифеста: пользовательское и служебное.
PROTECTED_NAMES = SKIP_NAMES | {
    MANIFEST_NAME,
    CLEANUP_MARKER,
    "version.json",
    "update_download",
    ROLLBACK_DIRNAME,
}

# Расширения заведомо программных файлов в корне папки (остальное в корне
# — возможные пользовательские файлы, их не трогаем).
PROGRAM_EXTS = {".dll", ".pyd", ".py", ".pyc", ".pyw", ".so", ".exe",
                ".zip", ".qm", ".qml", ".qmltypes"}


def _manifest_files(root: Path):
    """Читает манифест сборки. Возвращает (build, files, dirs).

    build — номер сборки из строки-заголовка (или None),
    files — множество путей файлов (маленькими буквами),
    dirs  — множество каталогов манифеста (все уровни).
    """
    mf = Path(root) / MANIFEST_NAME
    if not mf.is_file():
        return None, set(), set()
    build = None
    files, dirs = set(), set()
    try:
        for line in mf.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith("#"):
                m = re.search(r"build\s+(\d+)", line)
                if m:
                    build = int(m.group(1))
                continue
            low = line.lower()
            files.add(low)
            parts = low.split("/")
            for i in range(1, len(parts)):
                dirs.add("/".join(parts[:i]))
    except Exception:
        return None, set(), set()
    return build, files, dirs


def cleanup_stale_files(app_dir, log=None) -> list:
    """Убирает из папки программы файлы прошлых сборок.

    Новая версия, поставленная поверх старой (руками или обновлением),
    оставляет файлы, которых в новой сборке нет, — они и вычищаются по
    манифесту app_files.txt. Пользовательское не трогается: data/ с базами,
    pending_update, version.json, отчёты и любые файлы не из программных
    расширений в корне. Возвращает список удалённых относительных путей.
    """
    root = install_root(app_dir)
    build, files, dirs = _manifest_files(root)
    if not files:
        return []                                  # манифеста нет — не по чему сверять
    marker = root / CLEANUP_MARKER
    try:
        if build and marker.read_text(encoding="utf-8").strip() == str(build):
            return []                              # эта сборка уже наводила порядок
    except Exception:
        pass

    top_dirs = {d.split("/", 1)[0] for d in dirs if d}
    removed = []

    def _rm(path, rel):
        try:
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            else:
                path.unlink()
            removed.append(rel)
        except OSError:
            pass                                   # занят системой — оставим

    for dirpath, dirnames, filenames in os.walk(root):
        rel_dir = os.path.relpath(dirpath, root).replace(os.sep, "/")
        top = "" if rel_dir == "." else rel_dir.split("/", 1)[0]
        if rel_dir == ".":
            dirnames[:] = [d for d in dirnames if d.lower() not in PROTECTED_NAMES]
        elif top.lower() in top_dirs:
            pass                                   # дерево новой сборки — сверяем
        else:
            dirnames[:] = []                       # пользовательская папка — не ходим
            continue
        for name in filenames:
            rel = (rel_dir + "/" + name) if rel_dir != "." else name
            if rel.lower() in files:
                continue
            if rel_dir == ".":
                # в корне убираем только заведомо программные файлы
                if os.path.splitext(name)[1].lower() in PROGRAM_EXTS:
                    _rm(Path(dirpath) / name, rel)
            else:
                # внутри дерева новой сборки — всё принадлежит программе
                _rm(Path(dirpath) / name, rel)
        if rel_dir == ".":
            for name in list(dirnames):
                if name.lower() in top_dirs:
                    continue
                dpath = Path(dirpath) / name
                try:
                    children = list(dpath.iterdir())
                except OSError:
                    continue
                # корневой каталог, которого нет в новой сборке: убираем,
                # если выглядит python-пакетом (numpy, matplotlib и т.п.)
                if any(c.name == "__init__.py" or c.suffix == ".pyd" for c in children):
                    _rm(dpath, name)
                    if name in dirnames:
                        dirnames.remove(name)

    # пустые каталоги внутри программных деревьев новой сборки
    for d in sorted((p for p in root.rglob("*") if p.is_dir()),
                    key=lambda p: len(p.parts), reverse=True):
        rel = d.relative_to(root).as_posix()
        top = rel.split("/", 1)[0] if "/" in rel else ""
        if top and top.lower() in top_dirs:
            try:
                if not any(d.iterdir()):
                    d.rmdir()
                    removed.append(rel)
            except OSError:
                pass

    if build is not None:
        try:
            marker.write_text(str(build), encoding="utf-8")
        except Exception:
            pass
    if removed and log:
        try:
            log("убрано файлов прошлой сборки: %d" % len(removed))
        except Exception:
            pass
    return removed

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


def scan_version(version: str, build: int = 0) -> str:
    """
    Строка, которую понимает старая is_newer без номера сборки.
    2.0.0-ALPHA.20 + сборка 70 → 2.0.0-ALPHA.70
    На экране по-прежнему честная version (20).
    """
    ver = (version or "").strip()
    bld = int(build or 0)
    if not ver or bld <= 0:
        return ver
    core, sep, tail = ver.partition("-")
    if not sep or not tail:
        return ver
    name, _, _num = tail.partition(".")
    if not name:
        return ver
    return f"{core}-{name}.{bld}"


def is_newer(candidate: str, current: str, candidate_build: int = 0, current_build: int = 0) -> bool:
    """candidate новее current?

    Номер сборки — главный аргумент: сборки нумеруются сквозняком по всей
    истории программы, поэтому когда обе известны, решает сборка (имена
    версий вроде «BETA.1» формату «2.0.0-ALPHA.N» не подчиняются).
    Без сборок сравниваются строки версии по разрядам.
    """
    if not candidate:
        return False
    if not current:
        return True
    cb, vb = int(candidate_build or 0), int(current_build or 0)
    if cb > 0 and vb > 0:
        return cb > vb
    c, v = parse_version(candidate), parse_version(current)
    if c[0] != v[0]:
        return c[0] > v[0]
    if not c[1] and v[1]:
        return True
    if c[1] and not v[1]:
        return False
    if c[1] != v[1]:
        return c[1] > v[1]
    if cb > 0 or vb > 0:
        return cb > vb
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


def write_version_json(
    path: Path,
    version: str,
    build: int = 0,
    for_package: bool = False,
    display: str = "",
    url: str = "",
) -> None:
    """Пишет version.json. Ссылку на архив (url) не теряет:

    программа перезаписывает этот файл при запуске (для своих нужд),
    а ссылка нужна запасному пути обновления — поэтому url из уже
    существующего файла переносится в новый, если не задан явно.
    """
    ver = (version or "").strip()
    bld = int(build or 0)
    shown = (display or "").strip()
    if for_package:
        scan = scan_version(ver, bld)
        payload = {"name": "OVERTIMETAB", "version": scan}
        if bld > 0:
            payload["build"] = bld
        honest = shown or ver
        if honest and honest != scan:
            payload["display"] = honest
    else:
        payload = {"name": "OVERTIMETAB", "version": shown or ver}
        if bld > 0:
            payload["build"] = bld
    link = (url or "").strip()
    if not link:
        try:
            link = str(json.loads(path.read_text(encoding="utf-8")).get("url") or "").strip()
        except Exception:
            link = ""
    if link:
        payload["url"] = link
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def read_build_json(path: Path) -> int:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return int(data.get("build") or 0)
    except Exception:
        return 0


def format_version_label(version: str, build: int = 0) -> str:
    ver = (version or "").strip()
    if int(build or 0) > 0:
        # журнал плоский, имени версии у записи может не быть — тогда
        # заголовок блока просто «Сборка N»
        return f"{ver} · сборка {int(build)}" if ver else f"Сборка {int(build)}"
    return ver


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


def current_app_build(app_dir: Path) -> int:
    here = Path(__file__).resolve().parent
    root = install_root(app_dir)
    for p in (
        root / VERSION_JSON,
        root / "_internal" / VERSION_JSON,
        Path(app_dir) / VERSION_JSON,
        here / VERSION_JSON,
    ):
        if p.exists():
            b = read_build_json(p)
            if b:
                return b
    for p in (here / THEME_REL, Path(app_dir) / THEME_REL):
        if p.exists():
            b = _read_build_from_theme_file(p)
            if b:
                return b
    return 0


def _read_build_from_theme_file(path: Path) -> int:
    try:
        m = re.search(r"appBuild:\s*(\d+)", path.read_text(encoding="utf-8"))
        return int(m.group(1)) if m else 0
    except Exception:
        return 0


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


def _read_json_field_from_bytes(raw: bytes, key: str):
    if not raw:
        return None
    try:
        data = json.loads(raw.decode("utf-8", errors="replace"))
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    return data.get(key)


def display_of_package(package: Path) -> str:
    """Честный номер с экрана (поле display). Пусто — значит version уже честный."""
    package = Path(package)
    if package.is_file() and package.suffix.lower() == ".zip":
        try:
            with zipfile.ZipFile(package, "r") as zf:
                for n in zf.namelist():
                    if Path(n).name.lower() != "version.json":
                        continue
                    if "/data/" in ("/" + n.replace("\\", "/").lower()):
                        continue
                    val = _read_json_field_from_bytes(zf.read(n), "display")
                    return str(val or "").strip()
        except Exception:
            return ""
        return ""
    if package.is_dir():
        for cand in (package / VERSION_JSON, package / "_internal" / VERSION_JSON):
            if not cand.exists():
                continue
            try:
                data = json.loads(cand.read_text(encoding="utf-8"))
                val = str(data.get("display") or "").strip()
                if val:
                    return val
            except Exception:
                continue
    return ""


def build_of_package(package: Path) -> int:
    """Номер сборки из version.json внутри посылки."""
    package = Path(package)
    if package.is_file() and package.suffix.lower() == ".zip":
        try:
            with zipfile.ZipFile(package, "r") as zf:
                for n in zf.namelist():
                    if Path(n).name.lower() != "version.json":
                        continue
                    if "/data/" in ("/" + n.replace("\\", "/").lower()):
                        continue
                    return _extract_build_from_bytes(zf.read(n))
        except Exception:
            return 0
        return 0
    if package.is_dir():
        for cand in (package / VERSION_JSON, package / "_internal" / VERSION_JSON):
            if cand.exists():
                b = read_build_json(cand)
                if b:
                    return b
    return 0


def _extract_build_from_bytes(raw: bytes) -> int:
    if not raw:
        return 0
    try:
        data = json.loads(raw.decode("utf-8", errors="replace"))
        return int(data.get("build") or 0)
    except Exception:
        return 0


def describe_package(package: Path) -> dict:
    root = find_package_root(package)
    if not root:
        raise UpdateError("Это не папка и не архив OVERTIMETAB")
    ver = version_of_package(root)
    bld = build_of_package(root)
    shown = display_of_package(root)
    return {
        "root": str(root),
        "version": ver,
        "build": bld,
        "display": shown,
        "is_zip": root.is_file(),
    }


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
    bld = build_of_package(root)
    shown = display_of_package(root)
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

        if ver:
            write_version_json(tmp / VERSION_JSON, ver, bld, for_package=True, display=shown)
        meta = {"version": ver, "build": bld, "display": shown, "source": str(root)}
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
    bld = 0
    shown = ""
    meta = dest / META_NAME
    if meta.exists():
        try:
            data = json.loads(meta.read_text(encoding="utf-8"))
            ver = str(data.get("version") or "")
            bld = int(data.get("build") or 0)
            shown = str(data.get("display") or "")
        except Exception:
            ver = ""
            bld = 0
            shown = ""
    if not ver:
        ver = version_of_package(dest)
    if not bld:
        bld = build_of_package(dest)
    if not shown:
        shown = display_of_package(dest)
    return {"root": str(dest), "version": ver, "build": bld, "display": shown}


def cleanup_pending(app_dir: Path) -> None:
    dest = pending_dir(app_dir)
    if dest.exists():
        shutil.rmtree(dest, ignore_errors=True)


def cleanup_obsolete_zips(app_dir: Path, current_version: str, current_build: int = 0) -> list[str]:
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
                bld = build_of_package(child)
                if is_newer(ver, current_version, bld, current_build):
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


def pick_best_update(app_dir: Path, current_version: str, current_build: int = 0) -> Optional[dict]:
    """Самая новая посылка, которая новее текущей. Без версии — пропускаем при автопоиске."""
    best = None
    best_ver = current_version
    best_bld = current_build
    for src in scan_update_sources(app_dir):
        try:
            if src.resolve() == pending_dir(app_dir).resolve():
                info = staged_info(app_dir)
                if not info:
                    continue
                ver = info["version"]
                bld = int(info.get("build") or 0)
                if ver and is_newer(ver, current_version, bld, current_build):
                    return {
                        "root": info["root"],
                        "version": ver,
                        "build": bld,
                        "display": info.get("display") or "",
                        "already_staged": True,
                    }
                continue
            ver = version_of_package(src)
            if not ver:
                continue
            bld = build_of_package(src)
            if is_newer(ver, best_ver if best else current_version, bld, best_bld if best else current_build):
                best = {
                    "root": str(src),
                    "version": ver,
                    "build": bld,
                    "display": display_of_package(src),
                    "already_staged": False,
                }
                best_ver = ver
                best_bld = bld
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


def rollback_root() -> Path:
    """Documents\\OverTimeTab\\rollback — один слот предыдущей версии."""
    home = (os.environ.get("OVERTIMETAB_RECOVERY_HOME")
            or os.environ.get("USERPROFILE")
            or os.environ.get("HOME"))
    base = Path(home) / "Documents" if home else Path.home() / "Documents"
    return base / "OverTimeTab" / ROLLBACK_DIRNAME


def make_rollback_point(app_dir: Path) -> Optional[dict]:
    """
    Снимок текущей установки → Documents\\OverTimeTab\\rollback.
    Его читает recovery.py: после аварии можно откатиться на эту версию.
    Пользовательские данные (data/, базы) в снимок не входят.
    Возвращает meta-словарь или None (неудача не мешает обновлению).
    """
    tmp = None
    try:
        root = install_root(app_dir)
        if not _exe_in(root):
            return None
        dest = rollback_root()
        tmp = dest.with_name(dest.name + ".__tmp")
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.rmtree(tmp, ignore_errors=True)

        def _ignore(directory, names):
            skipped = []
            for n in names:
                if (n in SKIP_NAMES or n == "update_download"
                        or n == ROLLBACK_DIRNAME
                        or n.endswith(".sqlite") or n.endswith(".sqlite-wal")
                        or n.endswith(".sqlite-shm")):
                    skipped.append(n)
            return set(skipped)

        shutil.copytree(root, tmp, ignore=_ignore)
        ver = current_app_version(root)
        bld = current_app_build(root)
        meta = {
            "version": ver,
            "build": bld,
            "display": format_version_label(ver, bld),
            "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        (tmp / ROLLBACK_META_NAME).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        shutil.rmtree(dest, ignore_errors=True)
        tmp.rename(dest)
        return meta
    except Exception:
        if tmp is not None:
            shutil.rmtree(tmp, ignore_errors=True)
        return None


def launch_file_swap(source: Path, dest: Path, pid: int, backup: bool = True) -> Path:
    """Пишет скрытый .vbs в TEMP и запускает его без консоли.

    backup=True (по умолчанию) — перед переодеванием делает снимок
    текущей версии в rollback, чтобы recovery мог откатиться.
    """
    if backup:
        try:
            make_rollback_point(dest)
        except Exception:
            pass
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

    try:
        make_rollback_point(dest)
    except Exception:
        pass

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
            rest = (m.group(2) or "").strip()
            bm = re.search(r"сборка\s+(\d+)", rest, flags=re.IGNORECASE)
            build_num = int(bm.group(1)) if bm else 0
            date = rest
            if bm:
                date = re.sub(r"сборка\s+\d+\s*[—–-]?\s*", "", rest, flags=re.IGNORECASE).strip(" —–-")
            current = {
                "version": m.group(1).strip(),
                "build_num": build_num,
                "date": date,
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


def split_version_key(raw: str) -> tuple[str, int]:
    s = (raw or "").strip()
    if "+" in s:
        ver, _, tail = s.partition("+")
        try:
            return ver, int(tail)
        except ValueError:
            return ver, 0
    return s, 0


def changelog_since(text: str, from_version: str, to_version: str, to_build: int = 0) -> list[dict]:
    """
    Записи новее from_version, не новее to_version (+ сборка).
    Если from_version пустой — только текущая (чтобы не вывалить всю историю).
    """
    from_ver, from_bld = split_version_key(from_version)
    out: list[dict] = []
    for b in parse_changelog(text):
        ver = (b.get("version") or "").strip()
        bld = int(b.get("build_num") or 0)
        if not ver:
            continue
        if to_version and is_newer(ver, to_version, bld, to_build):
            continue
        if from_version:
            if not is_newer(ver, from_ver, bld, from_bld):
                continue
        elif to_version and ver != to_version:
            continue
        elif to_build and bld and bld != to_build:
            continue
        if not (b["added"] or b["changed"] or b["fixed"] or b["removed"]):
            continue
        out.append(b)
    return out


def _merge_changelog_blocks(blocks: list[dict]) -> dict:
    """Склеивает несколько блоков одной версии в один список без повторов."""
    out = {
        "version": "",
        "build_num": 0,
        "date": "",
        "added": [],
        "changed": [],
        "fixed": [],
        "removed": [],
        "build": [],
    }
    seen = {key: set() for key in ("added", "changed", "fixed", "removed", "build")}
    for b in blocks:
        if not out["version"]:
            out["version"] = (b.get("version") or "").strip()
        for key in seen:
            for item in b.get(key) or []:
                if not item or item in seen[key]:
                    continue
                seen[key].add(item)
                out[key].append(item)
    return out


def changelog_for_version(text: str, version: str) -> list[dict]:
    """
    Один накопленный блок текущей версии — все сборки этой версии вместе.
    Новая версия в CHANGELOG начинается с пустого раздела и сюда не попадает.
    """
    ver = (version or "").strip()
    if not ver:
        return []
    hits = [b for b in parse_changelog(text) if (b.get("version") or "").strip() == ver]
    if not hits:
        return []
    merged = _merge_changelog_blocks(hits)
    if not (merged["added"] or merged["changed"] or merged["fixed"] or merged["removed"]):
        return []
    return [merged]


def _md_inline(text: str) -> str:
    """**жирный** из журнала -> <b>жирный</b>, остальное экранируется.

    Окно «Что нового» показывает текст как форматированный (RichText);
    без этой замены звёздочки markdown оставались бы в тексте как есть.
    """
    import html as _html
    parts = str(text or "").split("**")
    out = []
    for i, chunk in enumerate(parts):
        if i % 2 == 0:
            out.append(_html.escape(chunk))
        else:
            out.append("<b>" + _html.escape(chunk) + "</b>")
    return "".join(out)


def _bullets(items: list[str]) -> str:
    return "<br>".join("• " + _md_inline(x) for x in items if x)


def merge_changelog_blocks(blocks: list[dict]) -> dict:
    """Сводный блок из записей нескольких сборок.

    Окно «Что нового» показывает одно объединённое описание обновления:
    все записи интервала в четырёх разделах, без разбивки по сборкам.
    """
    merged = {"version": "", "build_num": 0, "date": "",
              "added": [], "changed": [], "fixed": [], "removed": [], "build": []}
    for b in blocks or []:
        for key in ("added", "changed", "fixed", "removed"):
            merged[key].extend(b.get(key) or [])
    return merged


def whats_new_qml(blocks: list[dict]) -> list[dict]:
    """Окно «Что нового»: одна порция на всё обновление разом."""
    return changelog_for_qml([merge_changelog_blocks(blocks)])


# Стандартная фраза для обновления без записей в журнале (все правки
# внутренние): окно «Что нового» не должно молчать после обновления
FALLBACK_WHATS_NEW_TEXT = "Мелкие улучшения и исправления"


def whats_new_fallback() -> list[dict]:
    """Порция «Что нового», когда обновлению нечего рассказать."""
    return changelog_for_qml([{"changed": [FALLBACK_WHATS_NEW_TEXT]}])


def changelog_for_qml(blocks: list[dict]) -> list[dict]:
    """Плоские словари для QML Repeater."""
    out = []
    for b in blocks:
        added = list(b.get("added") or [])
        changed = list(b.get("changed") or [])
        fixed = list(b.get("fixed") or [])
        removed = list(b.get("removed") or [])
        out.append({
            "version": format_version_label(b.get("version") or "", int(b.get("build_num") or 0)),
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



_BUILD_TAG_RE = re.compile(r"<!--\s*b\s*:\s*(\d+)\s*-->\s*$")


def strip_build_tags(text: str) -> str:
    """Убирает невидимые пометки сборок (<!--b:239-->) из текста журнала."""
    return re.sub(r"<!--\s*b\s*:\s*\d+\s*-->", "", text or "")


def build_from_version_key(key: str) -> int:
    """Номер сборки из ключа настроек «ВЕРСИЯ+сборка» («BETA.1+239» → 239)."""
    m = re.search(r"\+(\d+)\s*$", str(key or ""))
    return int(m.group(1)) if m else 0


def parse_changelog_bullets(text: str) -> list[dict]:
    """Плоский список записей журнала с пометками сборок.

    Каждая запись: {build, section, text, version}. build — номер сборки
    из невидимой пометки в конце строки (<!--b:239-->); без пометки — 0
    (записи, появившиеся до введения пометок).
    """
    out: list[dict] = []
    version, section = "", None
    for raw_line in (text or "").splitlines():
        line = raw_line.rstrip()
        m = _CHANGELOG_HEADER_RE.match(line)
        if m:
            version = m.group(1).strip()
            section = None
            continue
        sm = _CHANGELOG_SECTION_RE.match(line)
        if sm:
            section = _CHANGELOG_SECTION_MAP.get(sm.group(1).strip().lower())
            continue
        if not section:
            continue
        if line.startswith("- "):
            body = line[2:].strip()
            bm = _BUILD_TAG_RE.search(body)
            build = int(bm.group(1)) if bm else 0
            body = _BUILD_TAG_RE.sub("", body).strip()
            out.append({"build": build, "section": section,
                        "text": body, "version": version})
        elif line.startswith("  ") and out:
            out[-1]["text"] += " " + line.strip()
    return out


def changelog_for_builds(text: str, since_build: int, upto_build: int) -> list[dict]:
    """Блоки «Что нового» для сборок в интервале (since_build, upto_build].

    Одна сборка — один блок, сверху новее. Так после обновления видно
    только то, что появилось с прошлого запуска, а перепрыг через
    несколько сборок (230 → 239) показывает всё накопившееся: 231–239.
    Записи без пометки сборки (появившиеся до её введения) не показываются.
    """
    since_build, upto_build = int(since_build or 0), int(upto_build or 0)
    bullets = [b for b in parse_changelog_bullets(text)
               if b["build"] > 0 and since_build < b["build"] <= upto_build]
    if not bullets:
        return []
    blocks = []
    for build in sorted({b["build"] for b in bullets}, reverse=True):
        group = [b for b in bullets if b["build"] == build]
        block = {"version": group[0]["version"], "build_num": build, "date": "",
                 "added": [], "changed": [], "fixed": [], "removed": [],
                 "build": []}
        for b in group:
            block[b["section"]].append(b["text"])
        blocks.append(block)
    return blocks

# ---------------------------------------------------------------------------
# Сетевое обновление (веб-хранилище). Полная закачка zip.
# ---------------------------------------------------------------------------

def _update_info_url(base_url: str) -> str:
    """Куда стучаться за version.json: либо сам файл, либо <url>/version.json."""
    base = (base_url or "").strip().rstrip("/")
    if base.lower().endswith("/version.json"):
        return base
    return base + "/version.json"


def parse_github_release_url(url: str) -> Optional[dict]:
    """
    Понял ли это GitHub-ссылка на релиз:
      github.com/<owner>/<repo>                          → latest
      github.com/<owner>/<repo>/releases                 → latest
      github.com/<owner>/<repo>/releases/latest          → latest
      github.com/<owner>/<repo>/releases/tag/<tag>
      github.com/<owner>/<repo>/releases/download/<tag>[/<asset>]
    Возвращает {"owner","repo","tag"} (tag "" = latest) или None.
    """
    u = (url or "").strip().rstrip("/")
    m = re.match(r"^https?://github\.com/([^/]+)/([^/]+)(/releases)?(?:/(.*))?$", u, re.IGNORECASE)
    if m:
        owner, repo = m.group(1), m.group(2)
        rest = (m.group(4) or "").strip().rstrip("/")
        if not rest or rest == "latest":
            return {"owner": owner, "repo": repo, "tag": ""}
        if rest.startswith("tag/"):
            tag = rest[4:].split("/")[0].strip()
            return {"owner": owner, "repo": repo, "tag": tag} if tag else None
        if rest.startswith("download/"):
            parts = rest.split("/")
            tag = parts[1].strip() if len(parts) > 1 else ""
            return {"owner": owner, "repo": repo, "tag": tag} if tag else None
        return None
    return None


def _github_display_from_title(title: str) -> str:
    """Отображаемое имя версии из заголовка релиза.

    Заголовок релиза выглядит как «OVERTIMETAB BETA.1 · сборка 217» —
    для уведомлений нужно именно отображаемое имя («BETA.1»), а не
    машинная версия из тега (2.0.0-ALPHA.…).
    """
    m = re.search(r"OVERTIMETAB\s+(.+?)\s*·", title or "", flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _github_build_from_title(title: str) -> int:
    """Сборку берём из заголовка релиза («… · сборка 126»)."""
    m = re.search(r"сборка\s+(\d+)", title or "", flags=re.IGNORECASE)
    return int(m.group(1)) if m else 0


def fetch_github_release_info(gh: dict, timeout: float = 15.0) -> tuple[Optional[dict], str]:
    """Достаёт актуальную версию через GitHub API. Возвращает (info, ошибка)."""
    owner = gh["owner"]
    repo = gh["repo"]
    tag = gh.get("tag") or ""
    if tag:
        api = f"https://api.github.com/repos/{owner}/{repo}/releases/tags/{urllib.parse.quote(tag)}"
    else:
        api = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    req = urllib.request.Request(
        api,
        headers={"Accept": "application/vnd.github+json", "User-Agent": "OVERTIMETAB-updater"},
    )
    try:
        with _urlopen_https(req, timeout) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="replace"))
    except urllib.error.HTTPError as e:
        code = e.code
        if code == 403:
            return None, "GitHub API временно ограничил запросы (лимит). Повторите позже."
        if code == 404:
            return None, "Такой релиз или тег на GitHub не найден."
        return None, f"GitHub API ответил ошибкой {code}."
    except Exception as e:
        return None, f"Не удалось связаться с GitHub API ({e})."
    if not isinstance(data, dict):
        return None, "GitHub вернул некорректные данные."
    if data.get("draft"):
        return None, "Релиз помечен как черновик — обновление недоступно."
    tag_name = str(data.get("tag_name") or "").strip()
    if not tag_name:
        return None, "На GitHub не найден тег релиза."
    version = tag_name[1:] if tag_name.startswith("v") else tag_name
    build = _github_build_from_title(str(data.get("name") or ""))
    zip_url = ""
    for asset in (data.get("assets") or []):
        name = str(asset.get("name") or "").lower()
        if name.endswith(".zip"):
            zip_url = str(asset.get("browser_download_url") or "")
            break
    if not zip_url:
        return None, "На релизе не найден zip-архив с программой."
    info = {"name": "OVERTIMETAB", "version": version, "build": build, "url": zip_url}
    display = _github_display_from_title(str(data.get("name") or ""))
    if display:
        info["display"] = display
    return info, ""


def _der_to_pem(der: bytes) -> str:
    """DER-сертификат из хранилища Windows -> PEM для load_verify_locations."""
    import base64
    return ("-----BEGIN CERTIFICATE-----\n"
            + base64.encodebytes(der).decode("ascii")
            + "-----END CERTIFICATE-----\n")


_WINDOWS_SSL_CONTEXT = None


def windows_ssl_context():
    """SSL-контекст со всеми сертификатами из хранилищ Windows.

    Обычный контекст urllib собирает доверенные корни с фильтром по
    назначению сертификата, из-за чего корпоративные центры сертификации
    (проверка трафика, внутренние УЦ) иногда не попадают в список —
    браузер при этом работает. Здесь загружаем всё из «Доверенных
    корневых» и «Промежуточных» хранилищ текущего пользователя и
    компьютера — так же, как браузер.
    """
    global _WINDOWS_SSL_CONTEXT
    if _WINDOWS_SSL_CONTEXT is not None:
        return _WINDOWS_SSL_CONTEXT
    import ssl
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    ctx.check_hostname = True
    for store in ("ROOT", "CA"):
        try:
            for cert, _enc, _props in ssl.enum_certificates(store):
                try:
                    ctx.load_verify_locations(cadata=_der_to_pem(cert))
                except Exception:
                    pass
        except Exception:
            pass  # не Windows — хранилищ нет, контекст останется пустым
    _WINDOWS_SSL_CONTEXT = ctx
    return ctx


# ── Проверка сертификатов силами самого Windows (WinHTTP) ──────────────
# Третья попытка для HTTPS: соединение и проверку сертификата делает сам
# Windows через WinHTTP — тот же механизм, которым пользуются Edge/Chrome:
# хранилища доверия и компьютера, и пользователя, автозагрузка недостающих
# промежуточных звеньев цепочки по ссылке из сертификата. Python про
# корпоративный центр может не знать, а Windows — знает.

# Папка для файла-отчёта о непринятом сертификате (Main задаёт папку
# данных — Program Files обычно недоступна для записи). Не задана —
# пишем рядом с программой, если можно, иначе во временную папку.
_SSL_REPORT_DIR = None
_LAST_SSL_REPORT = ""
_SSL_REPORT_NAME = "update_ssl_report.txt"


def set_ssl_report_dir(path) -> None:
    """Куда класть отчёт о непринятом сертификате (вызывает Main)."""
    global _SSL_REPORT_DIR
    _SSL_REPORT_DIR = Path(path) if path else None


def _ssl_report_path() -> Path:
    if _SSL_REPORT_DIR is not None:
        return Path(_SSL_REPORT_DIR) / _SSL_REPORT_NAME
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path.cwd()
    try:
        probe = base / (".probe_" + str(os.getpid()))
        probe.touch()
        probe.unlink()
        return base / _SSL_REPORT_NAME
    except Exception:
        return Path(tempfile.gettempdir()) / _SSL_REPORT_NAME


def _n_or_q(n) -> str:
    """Число для отчёта: «?», если посчитать не удалось."""
    return "?" if n is None else str(n)


def _machine_store_count(store: str):
    """Число сертификатов в хранилище ЛОКАЛЬНОГО КОМПЬЮТЕРА (Crypt32).

    Python видит только хранилища текущего пользователя, а корпоративные
    корни часто кладут групповыми политиками именно в хранилище
    компьютера — тогда Python про них не знает, а Windows и браузеры
    знают. Не получилось посчитать — None (в отчёте будет «?»).
    """
    import ctypes
    try:
        crypt = ctypes.WinDLL("crypt32.dll", use_last_error=True)
        crypt.CertOpenStore.restype = ctypes.c_void_p
        crypt.CertOpenStore.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                        ctypes.c_void_p, ctypes.c_uint,
                                        ctypes.c_wchar_p]
        crypt.CertEnumCertificatesInStore.restype = ctypes.c_void_p
        crypt.CertEnumCertificatesInStore.argtypes = [ctypes.c_void_p,
                                                      ctypes.c_void_p]
        crypt.CertCloseStore.restype = ctypes.c_int
        crypt.CertCloseStore.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        CERT_STORE_PROV_SYSTEM_W = 11                 # системное хранилище (Unicode)
        CERT_SYSTEM_STORE_LOCAL_MACHINE = 0x00020000  # хранилище компьютера
        CERT_STORE_READONLY_FLAG = 0x00008000
        handle = crypt.CertOpenStore(
            ctypes.c_void_p(CERT_STORE_PROV_SYSTEM_W), 0, None,
            CERT_SYSTEM_STORE_LOCAL_MACHINE | CERT_STORE_READONLY_FLAG,
            "Root" if store == "ROOT" else "CA")
        if not handle:
            return None
        try:
            count, ctx = 0, None
            while True:
                ctx = crypt.CertEnumCertificatesInStore(handle, ctx)
                if not ctx:
                    return count
                count += 1
        finally:
            crypt.CertCloseStore(handle, 0)
    except Exception:
        return None


def _windows_store_counts() -> dict:
    """Число сертификатов по хранилищам: (пользователь, компьютер)."""
    import ssl
    out = {}
    for store in ("ROOT", "CA"):
        user_n = None
        try:
            user_n = sum(1 for _ in ssl.enum_certificates(store))
        except Exception:
            pass
        out[store] = (user_n, _machine_store_count(store))
    if all(v == (None, None) for v in out.values()):
        return {}          # не Windows — хранилищ нет
    return out


def _write_ssl_report(url: str, attempts) -> str:
    """Текстовый отчёт для диагностики: что пробовали и почему не приняли.

    Пишется, когда сертификат сайта не приняли ВСЕ способы соединения.
    Возвращает путь файла (или пустую строку, если записать не удалось).
    """
    global _LAST_SSL_REPORT
    import datetime
    import ssl
    try:
        path = _ssl_report_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        lines = [
            "Отчёт о проверке обновлений (не принят сертификат сайта)",
            "Дата: " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Адрес: " + url,
            "Python " + sys.version.split()[0] + " · " + ssl.OPENSSL_VERSION,
            "",
            "Все способы соединения не приняли сертификат:",
        ]
        for name, exc in attempts:
            reason = getattr(exc, "reason", None) or exc
            lines.append("- " + name + ": " + str(reason).strip())
        try:
            counts = _windows_store_counts()
        except Exception:
            counts = {}
        if counts:
            parts = []
            for store in ("ROOT", "CA"):
                user_n, machine_n = counts.get(store, (None, None))
                machine_text = _n_or_q(machine_n)
                if machine_n == 0:
                    machine_text += " (или нет доступа)"
                parts.append(store + ": пользователь " + _n_or_q(user_n)
                             + ", компьютер " + machine_text)
            lines += ["", "Сертификатов в хранилищах Windows — "
                      + "; ".join(parts)]
        else:
            lines += ["", "Хранилища сертификатов Windows недоступны (не Windows)"]
        path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
        _LAST_SSL_REPORT = str(path)
    except Exception:
        _LAST_SSL_REPORT = ""
    return _LAST_SSL_REPORT


class _BytesResponse:
    """Ответ в памяти для пути WinHTTP: тот же интерфейс, что у urlopen."""

    def __init__(self, body: bytes, status: int = 200, headers=None):
        self._bio = io.BytesIO(body)
        self.status = status
        self.code = status
        self.headers = headers or {}

    def read(self, n=-1):
        return self._bio.read(n)

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


# Биты отказа WinHTTP — WINHTTP_CALLBACK_STATUS_FLAG_* из winhttp.h
# (точные значения из настоящего заголовка; по ним видно, ЧТО именно
# не понравилось Windows — отзыв, имя, срок, корень)
_WINHTTP_SECURE_BITS = [
    (0x00000001, "не удалось проверить отзыв сертификата (сервер списков отзыва недоступен)"),
    (0x00000002, "сертификат сайта недействителен"),
    (0x00000004, "сертификат сайта отозван"),
    (0x00000008, "центр сертификации сайта не входит в доверенные Windows"),
    (0x00000010, "имя в сертификате не совпадает с адресом сайта"),
    (0x00000020, "срок действия сертификата истёк или ещё не наступил"),
    (0x00000040, "сертификат не предназначен для сайтов"),
    (0x80000000, "внутренняя ошибка защищённого канала Windows"),
]


def _winhttp_flags_text(value: int) -> str:
    """Расшифровка кода, которым Windows отказал в доверии сертификату."""
    bits = [text for mask, text in _WINHTTP_SECURE_BITS if value & mask]
    return "; ".join(bits) if bits else "код отказа 0x%08X" % value


# Подробности отказа сертификата Windows сообщает только в callback-функцию
# (WINHTTP_CALLBACK_STATUS_SECURE_FAILURE, флаги — DWORD по указателю),
# поэтому держим ссылку на callback и место, куда он записывает флаги
_WINHTTP_SECURE_INFO = {"flags": None}
_WINHTTP_CALLBACK_REF = None


def _winhttp_status_callback(hreq, ctx, status, info, length):
    """Принимает от Windows подробности отказа сертификата (если были)."""
    if status == 0x00010000 and info:      # WINHTTP_CALLBACK_STATUS_SECURE_FAILURE
        try:
            import ctypes
            _WINHTTP_SECURE_INFO["flags"] = int(
                ctypes.cast(info, ctypes.POINTER(ctypes.c_ulong))[0])
        except Exception:
            pass


# Код последней ошибки WinHTTP: на Windows — настоящий из ctypes, в тестах
# без Windows — подставной (путь отказа сертификата проверяется везде)
_WINHTTP_TEST_LAST_ERROR = None


def _winhttp_last_error() -> int:
    import ctypes
    get = getattr(ctypes, "get_last_error", None)
    if get is not None:
        return int(get())
    return int(_WINHTTP_TEST_LAST_ERROR or 0)


def _winhttp_check(win, hreq) -> None:
    """Сбой запроса WinHTTP → понятное исключение (12175 — сертификат)."""
    import ctypes
    import ssl
    code = _winhttp_last_error()
    if code == 12175:  # ERROR_WINHTTP_SECURE_FAILURE
        flags = _WINHTTP_SECURE_INFO.get("flags")
        if flags is None:
            # запасной путь: WINHTTP_OPTION_SECURITY_FLAGS = 31
            val = ctypes.c_ulong(0)
            size = ctypes.c_ulong(4)
            if win.WinHttpQueryOption(hreq, 31, ctypes.byref(val),
                                      ctypes.byref(size)):
                flags = val.value
        detail = _winhttp_flags_text(flags) if flags else "причина не уточнена"
        raise ssl.SSLError(
            "Windows сам проверил сертификат сайта и не принял его: " + detail)
    raise OSError("WinHTTP, код " + str(code))


def _winhttp_header_str(win, hreq, which: int) -> str:
    """Строковый заголовок ответа (which — код WINHTTP_QUERY_…)."""
    import ctypes
    size = ctypes.c_ulong(0)
    win.WinHttpQueryHeaders(hreq, which, None, None, ctypes.byref(size), None)
    n = size.value // 2
    if n <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(n)
    if win.WinHttpQueryHeaders(hreq, which, None, buf, ctypes.byref(size), None):
        return buf.value
    return ""


def _winhttp_read_all(win, hreq) -> bytes:
    import ctypes
    chunks = []
    avail = ctypes.c_ulong(0)
    while True:
        if not win.WinHttpQueryDataAvailable(hreq, ctypes.byref(avail)):
            raise OSError("WinHttpQueryDataAvailable, код "
                          + str(_winhttp_last_error()))
        if avail.value == 0:
            break
        buf = ctypes.create_string_buffer(avail.value)
        got = ctypes.c_ulong(0)
        if not win.WinHttpReadData(hreq, buf, avail.value, ctypes.byref(got)):
            raise OSError("WinHttpReadData, код " + str(_winhttp_last_error()))
        if got.value:
            chunks.append(buf.raw[:got.value])
    return b"".join(chunks)


def _winhttp_prototypes():
    """Сигнатуры функций WinHTTP (restype, argtypes) — как в настоящем API.

    Число аргументов обязано совпадать с Windows точно: лишний или
    недостающий аргумент в ctypes либо падает («takes N arguments»), либо
    ломает стек вызова. HINTERNET — указатель, поэтому restype у функций,
    возвращающих обработчик, — c_void_p (иначе на 64-битной Windows адрес
    обрезался бы до int).
    """
    import ctypes
    from ctypes import wintypes
    return {
        "WinHttpOpen": (ctypes.c_void_p,
                        [ctypes.c_wchar_p, ctypes.c_uint,
                         ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]),
        "WinHttpConnect": (ctypes.c_void_p,
                           [ctypes.c_void_p, ctypes.c_wchar_p,
                            ctypes.c_ushort, ctypes.c_uint]),
        "WinHttpOpenRequest": (ctypes.c_void_p,
                               [ctypes.c_void_p, ctypes.c_wchar_p,
                                ctypes.c_wchar_p, ctypes.c_wchar_p,
                                ctypes.c_wchar_p, ctypes.c_wchar_p,
                                ctypes.c_uint]),
        "WinHttpSetTimeouts": (ctypes.c_int,
                               [ctypes.c_void_p] + [ctypes.c_int] * 4),
        "WinHttpSendRequest": (ctypes.c_int,
                               [ctypes.c_void_p, ctypes.c_wchar_p,
                                ctypes.c_uint, ctypes.c_void_p,
                                ctypes.c_uint, ctypes.c_uint,
                                ctypes.c_void_p]),
        "WinHttpReceiveResponse": (ctypes.c_int,
                                   [ctypes.c_void_p, ctypes.c_void_p]),
        "WinHttpQueryHeaders": (ctypes.c_int,
                                [ctypes.c_void_p, ctypes.c_uint,
                                 ctypes.c_wchar_p, ctypes.c_void_p,
                                 ctypes.POINTER(wintypes.DWORD),
                                 ctypes.c_void_p]),
        "WinHttpQueryDataAvailable": (ctypes.c_int,
                                      [ctypes.c_void_p,
                                       ctypes.POINTER(wintypes.DWORD)]),
        "WinHttpReadData": (ctypes.c_int,
                            [ctypes.c_void_p, ctypes.c_void_p,
                             ctypes.c_uint,
                             ctypes.POINTER(wintypes.DWORD)]),
        "WinHttpQueryOption": (ctypes.c_int,
                               [ctypes.c_void_p, ctypes.c_uint,
                                ctypes.c_void_p,
                                ctypes.POINTER(wintypes.DWORD)]),
        "WinHttpSetStatusCallback": (ctypes.c_void_p,
                                     [ctypes.c_void_p,
                                      ctypes.CFUNCTYPE(
                                          None, ctypes.c_void_p,
                                          ctypes.c_size_t, ctypes.c_uint,
                                          ctypes.c_void_p, ctypes.c_uint),
                                      ctypes.c_uint, ctypes.c_size_t]),
        "WinHttpCloseHandle": (ctypes.c_int, [ctypes.c_void_p]),
    }


def _winhttp_get(url: str, timeout: float) -> _BytesResponse:
    """HTTPS GET силами самого Windows (WinHTTP).

    Проверку сертификата выполняет Windows — браузерным механизмом.
    Ошибки сертификата приходят как ssl.SSLError (с расшифровкой кода),
    HTTP-ошибки — как urllib.error.HTTPError, чтобы вызывающему коду
    было всё равно, каким путём пришёл ответ.
    """
    import ctypes
    from ctypes import wintypes

    parts = urllib.parse.urlsplit(url)
    if parts.scheme.lower() != "https" or not parts.hostname:
        raise ValueError("WinHTTP-путь работает только с https-адресами")
    if sys.platform != "win32" and not os.environ.get("OVERTIMETAB_TEST_WINHTTP"):
        raise RuntimeError("WinHTTP доступен только на Windows")

    try:
        win = ctypes.WinDLL("winhttp.dll", use_last_error=True)
    except (OSError, AttributeError) as e:
        # AttributeError — ctypes без WinDLL (не Windows)
        raise RuntimeError("Библиотека WinHTTP недоступна: " + str(e)) from e

    protos = _winhttp_prototypes()
    for _name, (_restype, _argtypes) in protos.items():
        _fn = getattr(win, _name)
        _fn.restype = _restype
        _fn.argtypes = _argtypes

    ms = max(1000, int(timeout * 1000))
    hsession = win.WinHttpOpen("OVERTIMETAB-updater", 0, None, None, 0)
    if not hsession:
        raise OSError("WinHttpOpen, код " + str(_winhttp_last_error()))
    try:
        win.WinHttpSetTimeouts(hsession, 0, ms, ms, ms)
        # Подробности отказа сертификата Windows сообщает в callback
        # (WINHTTP_CALLBACK_STATUS_SECURE_FAILURE); держим ссылку, чтобы
        # сборщик мусора не забрал функцию, пока идёт запрос
        global _WINHTTP_CALLBACK_REF
        _WINHTTP_SECURE_INFO["flags"] = None
        _WINHTTP_CALLBACK_REF = protos["WinHttpSetStatusCallback"][1][1](
            _winhttp_status_callback)
        win.WinHttpSetStatusCallback(hsession, _WINHTTP_CALLBACK_REF,
                                     0x00010000, 0)
        current = url
        for _hop in range(6):            # до пяти переадресаций
            cur = urllib.parse.urlsplit(current)
            host = cur.hostname
            port = cur.port or 443
            path = cur.path or "/"
            if cur.query:
                path += "?" + cur.query
            hconn = win.WinHttpConnect(hsession, host, port, 0)
            if not hconn:
                raise OSError("WinHttpConnect, код " + str(_winhttp_last_error()))
            try:
                # 0x00800000 — WINHTTP_FLAG_SECURE (https)
                hreq = win.WinHttpOpenRequest(hconn, "GET", path, None,
                                              None, None, 0x00800000)
                if not hreq:
                    raise OSError("WinHttpOpenRequest, код "
                                  + str(_winhttp_last_error()))
                try:
                    if not win.WinHttpSendRequest(hreq, None, 0, None, 0, 0, None):
                        _winhttp_check(win, hreq)
                    if not win.WinHttpReceiveResponse(hreq, None):
                        _winhttp_check(win, hreq)
                    status = wintypes.DWORD(0)
                    size = wintypes.DWORD(4)
                    if not win.WinHttpQueryHeaders(
                            hreq, 19 | 0x20000000, None, ctypes.byref(status),
                            ctypes.byref(size), None):
                        raise OSError("WinHttpQueryHeaders, код "
                                      + str(_winhttp_last_error()))
                    if status.value in (301, 302, 303, 307, 308):
                        loc = _winhttp_header_str(win, hreq, 33)  # Location
                        if not loc:
                            raise OSError("переадресация без адреса")
                        current = urllib.parse.urljoin(current, loc)
                        continue
                    body = _winhttp_read_all(win, hreq)
                    if status.value >= 400:
                        raise urllib.error.HTTPError(
                            current, status.value,
                            "HTTP " + str(status.value), None, io.BytesIO(body))
                    return _BytesResponse(
                        body, status.value,
                        {"Content-Length": str(len(body))})
                finally:
                    win.WinHttpCloseHandle(hreq)
            finally:
                win.WinHttpCloseHandle(hconn)
        raise RuntimeError("слишком много переадресаций")
    finally:
        win.WinHttpCloseHandle(hsession)


def _is_cert_error(exc: Exception) -> bool:
    """Ошибка доверия к сертификату сайта (CERTIFICATE_VERIFY_FAILED)."""
    import ssl
    if isinstance(exc, ssl.SSLError):
        return True
    reason = getattr(exc, "reason", None)
    return isinstance(reason, ssl.SSLError)


def _ssl_fallback_enabled() -> bool:
    # на Windows — всегда; переменная ниже оставлена для тестов не на Windows
    return sys.platform == "win32" or bool(
        os.environ.get("OVERTIMETAB_TEST_SSL_FALLBACK"))


def _urlopen_https(url, timeout: float, data=None):
    """HTTPS-запрос с запасными путями для корпоративных сертификатов.

    Порядок попыток:
      1. Стандартный контекст Python.
      2. Контекст со всеми корнями из хранилищ Windows (ROOT + CA).
      3. Соединение силами самого Windows (WinHTTP) — как браузеры:
         все хранилища доверия, автозагрузка недостающих звеньев цепочки.
    Если сертификат не приняли все три — пишем файл-отчёт для диагностики.
    """
    import ssl
    try:
        return urllib.request.urlopen(url, timeout=timeout, data=data)
    except (urllib.error.URLError, ssl.SSLError) as e1:
        if not (_is_cert_error(e1) and _ssl_fallback_enabled()):
            raise
        attempts = [("обычный SSL Python", e1)]
        try:
            return urllib.request.urlopen(
                url, timeout=timeout, data=data,
                context=windows_ssl_context())
        except (urllib.error.URLError, ssl.SSLError) as e2:
            if not _is_cert_error(e2):
                raise
            attempts.append(("корни из хранилищ Windows", e2))
            # WinHTTP — только простые GET-запросы по строке-адресу
            if data is None and isinstance(url, str) \
                    and url.lower().startswith("https://"):
                try:
                    return _winhttp_get(url, timeout)
                except ssl.SSLError as e3:
                    attempts.append(("проверка Windows (WinHTTP)", e3))
                    _write_ssl_report(url, attempts)
                    raise
                except Exception as e3:
                    # WinHTTP не справился по другой причине — показываем
                    # исходную ошибку сертификата, отчёт тоже пригодится
                    attempts.append(("проверка Windows (WinHTTP) не запустилась", e3))
                    _write_ssl_report(url, attempts)
                    raise e2 from e3
            _write_ssl_report(url, attempts)
            raise


def _cert_error_message(url: str, exc: Exception) -> str:
    """Понятное объяснение ошибки сертификата для окна проверки обновлений."""
    host = url.split("//", 1)[-1].split("/", 1)[0]
    msg = ("Компьютер не доверяет сертификату сайта " + host + " ("
           + str(getattr(exc, "reason", exc)) + "). Обычно причина — "
           "защита сетевого трафика на рабочем компьютере. Обратитесь "
           "к системному администратору.")
    if _LAST_SSL_REPORT:
        msg += "\nПодробности — в файле " + _LAST_SSL_REPORT
    return msg


def _fetch_version_json(base_url: str, timeout: float = 15.0) -> tuple[Optional[dict], str]:
    """Классический путь: читаем version.json с хранилища прямо в память."""
    url = _update_info_url(base_url)
    try:
        with _urlopen_https(url, timeout) as resp:
            raw = resp.read()
        data = json.loads(raw.decode("utf-8", errors="replace"))
    except Exception as e:
        if _is_cert_error(e):
            return None, _cert_error_message(url, e)
        return None, f"Не удалось прочитать version.json ({e})."
    if not isinstance(data, dict):
        return None, "version.json содержит некорректные данные."
    return data, ""


def _raw_github_version_json(gh: dict, timeout: float = 15.0) -> tuple[Optional[dict], str]:
    """
    Запасной путь, когда api.github.com недоступен: читает updates/version.json
    из репозитория по стабильному рефу (тег релиза → main) через
    raw.githubusercontent.com. Версия + имя актуального zip обновляются в файле
    после каждой сборки, поэтому ссылку на GitHub вставлять надо один раз —
    имя архива менять не придётся.
    Возвращает info c полем base_url (куда качать zip).
    """
    owner, repo = gh["owner"], gh["repo"]
    tag = gh.get("tag") or ""
    refs = []
    if tag:
        refs.append(tag)
    refs.append("main")
    refs.append("arena/01a043e7-overtimeapp")
    for ref in refs:
        url = f"https://raw.githubusercontent.com/{owner}/{repo}/{urllib.parse.quote(ref)}/updates/version.json"
        try:
            with _urlopen_https(url, timeout) as resp:
                raw = resp.read()
            data = json.loads(raw.decode("utf-8", errors="replace"))
        except Exception:
            continue
        if not isinstance(data, dict) or not data.get("version"):
            continue
        # url в version.json — ИМЯ архива (OVERTIMETAB_…_….zip), а не адрес:
        # файл читают разные хранилища (GitHub, запасной сервер), и имя
        # склеивается с адресом того, где файл лежит. Для «latest» тега нет,
        # поэтому пользуемся постоянным адресом последнего релиза
        # (/releases/latest/download/<архив>) — GitHub сам подставит тег.
        if tag:
            data["base_url"] = f"https://github.com/{owner}/{repo}/releases/download/{tag}"
        else:
            data["base_url"] = f"https://github.com/{owner}/{repo}/releases/latest/download"
        return data, ""
    return None, "Не удалось получить данные об обновлении (GitHub API недоступен, version.json в репозитории не найден)."


def fetch_update_info(base_url: str, timeout: float = 15.0) -> tuple[Optional[dict], str]:
    """
    Узнаёт про новую версию, ничего не скачивая на диск.
      Прямая ссылка на .zip → качаем сразу, версию узнаем из самого архива.
      GitHub-ссылка        → GitHub API; если недоступен — version.json из
                              репозитория (raw.githubusercontent.com), затем
                              version.json как ассет релиза.
      Любая другая         → читает version.json прямо в память.
    Возвращает (info, ошибка): info dict с version/build/url (+ base_url /
    direct_zip), ошибка — пусто при успехе.
    """
    base = (base_url or "").strip()
    if base.lower().endswith(".zip"):
        # Точный архив: version/build узнаем после скачивания (describe_package).
        return {"name": "OVERTIMETAB", "version": "", "build": 0, "url": base, "direct_zip": True}, ""
    gh = parse_github_release_url(base)
    if gh:
        info, err = fetch_github_release_info(gh, timeout)
        if info:
            return info, ""
        # API недоступен/лимит — стабильный version.json из репозитория.
        ri, rerr = _raw_github_version_json(gh, timeout)
        if ri:
            return ri, ""
        # Затем version.json рядом (ассет релиза).
        vi, verr = _fetch_version_json(base, timeout)
        if vi:
            return vi, ""
        return None, err or rerr or verr or "Не удалось получить данные об обновлении с GitHub."
    return _fetch_version_json(base, timeout)



def resolve_download_url(zip_ref: str, base_url: str) -> str:
    """Относительную ссылку на архив склеивает с базовым адресом, абсолютную оставляет."""
    ref = (zip_ref or "").strip()
    if re.match(r"^https?://", ref, re.IGNORECASE):
        return ref
    base = (base_url or "").strip().rstrip("/")
    return base + "/" + ref.lstrip("/")


def _fmt_mb(n: int) -> str:
    try:
        return f"{n / (1024 * 1024):.1f} МБ"
    except Exception:
        return str(n)


def download_zip(
    download_url: str,
    dest_dir: Path,
    sha256: str = "",
    timeout: float = 120.0,
    progress=None,
) -> Path:
    """
    Скачивает zip в dest_dir, возвращает путь к готовому файлу.
    progress(done:int, total:int) вызывается по мере загрузки (total может быть 0).
    Если задан sha256 — сверяет контрольную сумму и при несовпадении откатывает.
    """
    import hashlib

    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    name = download_url.split("?")[0].split("/")[-1] or "update.zip"
    if not name.lower().endswith(".zip"):
        name += ".zip"
    target = dest_dir / name
    tmp = dest_dir / (name + ".part")
    for p in (tmp, target):
        try:
            p.unlink(missing_ok=True)
        except Exception:
            pass
    try:
        with _urlopen_https(download_url, timeout) as resp:
            total = 0
            cl = resp.headers.get("Content-Length")
            if cl:
                try:
                    total = int(cl)
                except Exception:
                    total = 0
            done = 0
            with open(tmp, "wb") as out:
                while True:
                    chunk = resp.read(1 << 16)
                    if not chunk:
                        break
                    out.write(chunk)
                    done += len(chunk)
                    if progress:
                        try:
                            progress(done, total)
                        except Exception:
                            pass
        if sha256:
            h = hashlib.sha256()
            with open(tmp, "rb") as f:
                for chunk in iter(lambda: f.read(1 << 16), b""):
                    h.update(chunk)
            if h.hexdigest().lower() != sha256.lower():
                tmp.unlink(missing_ok=True)
                raise UpdateError("Контрольная сумма скачанного файла не совпала — файл подменён или битый")
        tmp.rename(target)
        return target
    except UpdateError:
        raise
    except Exception:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
        raise


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
