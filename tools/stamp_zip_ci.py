#!/usr/bin/env python3
"""Гарантия version.json в релизном zip (запускается в CI при сборке установщика).

Сборки 294–298 научили нас плохому: клиенты не находили в архиве номер
версии и отказывались обновляться. Каким бы ни было звено, из-за которого
файл не доехал (PyInstaller кладёт данные в _internal, и любой сбой
в этой цепочке оставлял архив без версии), архив обязан содержать
version.json В КОРНЕ — его читают ВСЕ поколения обновлятора, от самых
старых до текущих.

Скрипт вызывается из tools/overtimetab.iss на этапе компиляции
установщика (ISPP #exec): шаг «Собрать ZIP» в CI уже прошёл, а загрузка
релиза ещё нет — докладывание попадает в опубликованный архив.

Содержимое берётся из build/version.json (его пишет tools/overtimetab.spec);
если файла нет — вычисляется из AppTheme.qml тем же правилом (машинная
строка-мост: 2.0.0-<канал>.<сборка>).

Запуск: python tools/stamp_zip_ci.py [рабочая_папка]
(по умолчанию — корень репозитория; рабочая папка нужна для тестов).
"""
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def _identity() -> dict:
    """Машинная строка и сборка — из build/version.json или из AppTheme."""
    meta_path = ROOT / "build" / "version.json"
    try:
        data = json.loads(meta_path.read_text(encoding="utf-8"))
        ver = str(data.get("version") or "").strip()
        bld = int(data.get("build") or 0)
        if ver and bld > 0:
            out = {"name": "OVERTIMETAB", "version": ver, "build": bld}
            if str(data.get("display") or "").strip():
                out["display"] = str(data["display"]).strip()
            return out
    except Exception:
        pass
    # Запасной путь: вычисляем сами из AppTheme (то же правило, что в спеке)
    theme = ROOT / "components" / "AppTheme.qml"
    text = theme.read_text(encoding="utf-8") if theme.exists() else ""
    name_m = re.search(r'appVersion:\s*"([^"]+)"', text)
    name = name_m.group(1) if name_m else ""
    sem_m = re.search(r'appSemver:\s*"([^"]+)"', text)
    semver = sem_m.group(1) if sem_m else name
    bld_m = re.search(r"appBuild:\s*(\d+)", text)
    bld = int(bld_m.group(1)) if bld_m else 0
    sys.path.insert(0, str(ROOT))
    from app_update import scan_version  # noqa: E402
    out = {"name": "OVERTIMETAB", "version": scan_version(semver, bld), "build": bld}
    if name and name != out["version"]:
        out["display"] = name
    return out


def _norm(name: str) -> str:
    return name.replace("\\", "/").lstrip("/").lower()


def _meta_bytes(meta: dict) -> bytes:
    return (json.dumps(meta, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def find_zip(workspace: Path):
    """Свежий OVERTIMETAB_*.zip в рабочей папке (как его именует CI)."""
    zips = [p for p in workspace.glob("OVERTIMETAB_*.zip") if p.is_file()]
    if not zips:
        return None
    return max(zips, key=lambda p: p.stat().st_mtime)


def stamp(zip_path: Path, meta: dict) -> bool:
    """Докладывает version.json в корень и _internal, если их нет или они чужие.

    Возвращает True, если архив изменён.
    """
    payload = _meta_bytes(meta)
    want = str(meta.get("version") or "")
    with zipfile.ZipFile(zip_path, "r") as zf:
        names = zf.namelist()
        have = {_norm(n) for n in names}
        current = {}
        for target in ("version.json", "_internal/version.json"):
            if target in have:
                try:
                    current[target] = json.loads(
                        zf.read(target).decode("utf-8", errors="replace"))
                except Exception:
                    current[target] = None
    need = []
    for target in ("version.json", "_internal/version.json"):
        if target not in have:
            need.append(target)
        elif not isinstance(current.get(target), dict) or \
                str(current[target].get("version") or "") != want or \
                int(current[target].get("build") or 0) != int(meta.get("build") or 0):
            need.append(target)
    if not need:
        print(f"[zip-stamp] архив в порядке: version={want} build={meta.get('build')}")
        return False
    # Докладываем недостающее (простое добавление)
    add = [t for t in need if t not in have]
    if add:
        with zipfile.ZipFile(zip_path, "a") as zf:
            for target in add:
                zf.writestr(target, payload)
                print(f"[zip-stamp] добавлен {target}")
    # Чужое содержимое — пересобираем архив целиком (zip нельзя править на месте)
    replace = [t for t in need if t in have]
    if replace:
        tmp = zip_path.with_name(zip_path.name + ".__restamp")
        with zipfile.ZipFile(zip_path, "r") as src, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                if _norm(info.filename) in set(replace):
                    continue
                dst.writestr(info, src.read(info.filename))
            for target in replace:
                dst.writestr(target, payload)
                print(f"[zip-stamp] заменён {target}")
        tmp.replace(zip_path)
    print(f"[zip-stamp] версия гарантирована: version={want} build={meta.get('build')}")
    return True


def main() -> int:
    workspace = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else ROOT
    zip_path = find_zip(workspace)
    if not zip_path:
        # Не смерть: шаг «Собрать ZIP» мог поменяться — но громко предупреждаем
        print("[zip-stamp] WARNING: OVERTIMETAB_*.zip не найден — нечего проверять")
        return 0
    meta = _identity()
    if not meta.get("version") or not int(meta.get("build") or 0):
        print("[zip-stamp] ERROR: не смогли определить версию программы")
        return 1
    print(f"[zip-stamp] {zip_path.name}")
    try:
        stamp(zip_path, meta)
    except zipfile.BadZipFile:
        print("[zip-stamp] ERROR: архив битый — релиз нельзя публиковать")
        return 1
    except Exception as e:
        print(f"[zip-stamp] ERROR: {e}")
        return 1
    # Контрольная проверка: архив открывается и версия читается
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            data = json.loads(zf.read("version.json").decode("utf-8"))
        assert data.get("version") == meta["version"], data
        print("[zip-stamp] контроль: version.json в корне читается")
    except Exception as e:
        print(f"[zip-stamp] ERROR: контрольная проверка не прошла: {e}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
