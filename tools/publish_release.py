# -*- coding: utf-8 -*-
"""
Публикует GitHub Release из версии в AppTheme.qml и текста CHANGELOG.md.

Релиз состоит из ДВУХ обязательных файлов (иначе публикация не проходит):
    OVERTIMETAB_<версия>_<sha>.exe — установщик для человека;
    OVERTIMETAB_<версия>_<sha>.zip — обновление для установленной
        программы: автообновление ищет на релизе файл с .zip на конце.
Скрипт сам дожидается сборки Windows, проверяет ОБА файла, снимает
старые ассеты и проверяет, что сборка описана в BUILD_HISTORY.md.
Полный процесс выпуска — RELEASING.md.

Запуск из корня репозитория:
    python tools/publish_release.py
    python tools/publish_release.py --dry-run
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
THEME = ROOT / "components" / "AppTheme.qml"
CHANGELOG = ROOT / "CHANGELOG.md"


def read_identity() -> tuple[str, int]:
    text = THEME.read_text(encoding="utf-8")
    m = re.search(r'appVersion:\s*"([^"]+)"', text)
    if not m:
        sys.exit("Не нашли appVersion в AppTheme.qml")
    bm = re.search(r"appBuild:\s*(\d+)", text)
    return m.group(1), int(bm.group(1)) if bm else 0


def read_version() -> str:
    return read_identity()[0]


def changelog_section(version: str) -> str:
    if not CHANGELOG.exists():
        sys.exit("Нет CHANGELOG.md — сначала запишите, что изменилось")
    text = CHANGELOG.read_text(encoding="utf-8")
    lines = text.splitlines()
    start = None
    heading_re = re.compile(rf"^##\s+{re.escape(version)}\b")
    for i, line in enumerate(lines):
        if heading_re.search(line):
            start = i
            break
    if start is None:
        sys.exit(f"В CHANGELOG.md нет раздела «{version}»")
    end = len(lines)
    for j in range(start + 1, len(lines)):
        if lines[j].startswith("## "):
            end = j
            break
    chunk = lines[start:end]
    while chunk and chunk[-1].strip() in ("", "---"):
        chunk.pop()
    body = "\n".join(chunk).strip()
    # невидимые пометки сборок (<!--b:239-->) в текст релиза не идут
    body = re.sub(r"<!--\s*b\s*:\s*\d+\s*-->", "", body)
    return body + "\n"


def history_has_build(build: int) -> bool:
    """Описана ли сборка в BUILD_HISTORY.md (карта сборок из истории git).

    Карта — память проекта, независимая от ассистента: каждый, кто
    сопровождает табель (в том числе ИИ), знает, что въехало в какую
    сборку. Публикация без описания сборки запрещена.
    """
    hist = ROOT / "BUILD_HISTORY.md"
    if not hist.exists():
        return False
    try:
        return ("## Сборка %d" % int(build)) in hist.read_text(encoding="utf-8")
    except Exception:
        return False


def run_gh(args: list[str], check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["gh", *args],
        cwd=str(ROOT),
        text=True,
        capture_output=True,
        check=check,
    )


def current_branch() -> str:
    out = subprocess.check_output(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"],
        cwd=str(ROOT),
        text=True,
    )
    return out.strip()


def release_exists(tag: str) -> bool:
    r = run_gh(["release", "view", tag], check=False)
    return r.returncode == 0


def current_short_sha() -> str:
    out = subprocess.check_output(
        ["git", "rev-parse", "--short=7", "HEAD"],
        cwd=str(ROOT),
        text=True,
    )
    return out.strip()


def asset_errors(assets: list, keep_sha: str) -> list[str]:
    """Проверка состава релиза: zip и exe-установщик, оба с текущим sha.

    Релиз обязан содержать ДВА файла: установщик (*.exe) для человека и
    архив (*.zip) для установленной программы (автообновление ищет файл
    с .zip на конце — см. app_update.fetch_github_release_info).
    """
    keep = (keep_sha or "").lower()
    errors = []
    names = [str(a.get("name") or "") for a in (assets or [])]
    zips = [n for n in names if n.lower().endswith(".zip")]
    exes = [n for n in names if n.lower().endswith(".exe")]
    if not any(keep in n.lower() for n in zips):
        errors.append(
            "на релизе нет архива OVERTIMETAB_*_%s.zip — установленным "
            "программам нечем обновляться" % keep)
    if not any(keep in n.lower() for n in exes):
        errors.append(
            "на релизе нет установщика OVERTIMETAB_*_%s.exe — человеку "
            "нечего скачивать (см. RELEASING.md: релиз = zip + exe, "
            "сборкой управляет .github/workflows/build-windows.yml)" % keep)
    return errors


def release_assets(tag: str) -> list:
    view = run_gh(["release", "view", tag, "--json", "assets"], check=False)
    if view.returncode != 0:
        return []
    try:
        return json.loads(view.stdout or "{}").get("assets") or []
    except Exception:
        return []


def wait_for_build(timeout_sec: int = 1200) -> bool:
    """Ждёт завершения свежего запуска «Сборка Windows»."""
    import time as _time
    print("ждём сборку Windows на Actions…")
    _time.sleep(20)
    r = run_gh(["run", "list", "--workflow=build-windows.yml", "--limit", "1",
                "--json", "databaseId,status,conclusion"], check=False)
    try:
        info = json.loads(r.stdout or "[]")
        run_id = (info[0] or {}).get("databaseId")
    except Exception:
        run_id = None
    if not run_id:
        print("не удалось найти запуск сборки — проверьте вручную: "
              "gh run watch --workflow=build-windows.yml")
        return False
    watch = subprocess.run(["gh", "run", "watch", str(run_id), "--exit-status",
                            "--interval", "20"], cwd=str(ROOT))
    return watch.returncode == 0


def cleanup_old_release_assets(tag: str, keep_sha: str) -> None:
    """На одном теге — по одному zip и одному exe текущего sha.
    Actions именует файлы хешем коммита, новый не затирает старый —
    снимаем хвосты сами (и зипы, и установщики)."""
    keep = (keep_sha or "").lower()
    for asset in release_assets(tag):
        name = str(asset.get("name") or "")
        low = name.lower()
        if not (low.endswith(".zip") or low.endswith(".exe")):
            continue
        if keep and keep in low:
            continue
        gone = run_gh(["release", "delete-asset", tag, name, "--yes"], check=False)
        if gone.returncode == 0:
            print(f"сняли старый файл {name}")
        else:
            sys.stderr.write(gone.stderr or gone.stdout or f"не сняли {name}\n")


def main() -> int:
    parser = argparse.ArgumentParser(description="Опубликовать GitHub Release")
    parser.add_argument("--dry-run", action="store_true", help="только показать текст, не публиковать")
    parser.add_argument("--prerelease", action="store_true",
                        help="пометить релиз как пререлиз (по умолчанию — обычный релиз)")
    parser.add_argument("--no-wait", action="store_true",
                        help="не ждать сборку и не проверять ассеты (экстренные случаи)")
    parser.add_argument("--tag", default="",
                        help="тег существующего релиза (например v2.0.0-ALPHA.20), "
                             "если имя версии сменилось, а обновляться должны старые клиенты")
    args = parser.parse_args()

    version, build = read_identity()

    # Ссылка на архив в updates/version.json — ИМЯ файла (OVERTIMETAB_…_.zip),
    # а не полный адрес: файл читают разные хранилища (GitHub и запасной сервер
    # post.mvd.ru), и имя склеивается с адресом того, где файл лежит.
    # Полный адрес здесь запрещён сознательно: он привязал бы запасной путь
    # к GitHub, и при его недоступности качать было бы неоткуда.
    vj = ROOT / "updates" / "version.json"
    if vj.exists():
        try:
            vj_url = str(json.loads(vj.read_text(encoding="utf-8")).get("url", ""))
        except Exception as e:
            print(f"updates/version.json не читается: {e}", file=sys.stderr)
            return 1
        if not re.fullmatch(r"OVERTIMETAB_[^\s/]+\.zip", vj_url):
            print("В updates/version.json поле \"url\" должно быть ИМЕНЕМ архива "
                  "вида OVERTIMETAB_<версия>_<sha>.zip, а не адресом (сейчас: "
                  + repr(vj_url) + ").",
                  file=sys.stderr)
            return 1

    # Карта сборок: текущая сборка обязана быть описана в BUILD_HISTORY.md.
    # Файл генерируется из истории git: python tools/build_history.py --write --tag.
    if not history_has_build(build):
        print(f"В BUILD_HISTORY.md нет раздела \"Сборка {build}\". Запустите "
              "python tools/build_history.py --write --tag, закоммитьте файл "
              "и повторите публикацию.",
              file=sys.stderr)
        return 1

    tag = args.tag.strip() or f"v{version}"
    title = f"OVERTIMETAB {version}" + (f" · сборка {build}" if build else "")
    notes = changelog_section(version)
    branch = current_branch()
    pre = args.prerelease

    print(f"версия:   {version}" + (f" · сборка {build}" if build else ""))
    print(f"тег:      {tag}")
    print(f"ветка:    {branch}")
    print(f"пререлиз: {pre}")
    print("--- текст ---")
    print(notes, end="" if notes.endswith("\n") else "\n")
    print("---")

    if args.dry_run:
        print("dry-run: релиз не трогали")
        return 0

    notes_file = ROOT / ".release-notes.tmp.md"
    notes_file.write_text(notes, encoding="utf-8")
    try:
        if release_exists(tag):
            # Та же версия, новая сборка: тег переносим на этот коммит, zip потом заменит Actions.
            subprocess.run(["git", "tag", "-f", tag], cwd=str(ROOT), check=True)
            push = subprocess.run(
                ["git", "push", "origin", tag, "--force"],
                cwd=str(ROOT),
                text=True,
                capture_output=True,
            )
            if push.returncode != 0:
                sys.stderr.write(push.stderr or push.stdout or "не смогли сдвинуть тег\n")
                return push.returncode
            cmd = [
                "release", "edit", tag,
                "--title", title,
                "--notes-file", str(notes_file),
            ]
            cmd.append("--prerelease" if pre else "--latest")
            r = run_gh(cmd, check=False)
            action = "обновили"
        else:
            cmd = [
                "release", "create", tag,
                "--title", title,
                "--notes-file", str(notes_file),
                "--target", branch,
            ]
            if pre:
                cmd.append("--prerelease")
            else:
                cmd.append("--latest")
            r = run_gh(cmd, check=False)
            action = "создали"
        if r.returncode != 0:
            sys.stderr.write(r.stderr or r.stdout or "gh не смог опубликовать релиз\n")
            return r.returncode
        print((r.stdout or "").strip() or f"{action} релиз {tag}")
    finally:
        if notes_file.exists():
            notes_file.unlink()

    if args.no_wait:
        print("Запущено без ожидания (--no-wait): проверьте ассеты сами — "
              "на релизе обязаны быть zip и exe одной сборки.")
        return 0

    # ── Строгий контроль релиза ──────────────────────────────
    # Релиз = zip (обновление для программ) + exe (установщик для
    # человека). Ждём сборку и проверяем оба файла; старые снимаем.
    if not wait_for_build():
        print("Сборка Windows не завершилась успехом — релиз неполный.",
              file=sys.stderr)
        return 1
    sha = current_short_sha()
    assets = release_assets(tag)
    problems = asset_errors(assets, sha)
    if problems:
        for p in problems:
            print("ОШИБКА РЕЛИЗА: " + p, file=sys.stderr)
        return 1
    cleanup_old_release_assets(tag, sha)
    print(f"Релиз в порядке: zip + exe сборки {build} ({sha}).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
