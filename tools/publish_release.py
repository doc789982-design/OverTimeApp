# -*- coding: utf-8 -*-
"""
Публикует GitHub Release из версии в AppTheme.qml и текста CHANGELOG.md.

Я запускаю это после заметной версии. Тег v… на GitHub поднимает
workflow «Сборка Windows» — он сам приложит zip к этому же релизу.

Запуск из корня репозитория:
    python tools/publish_release.py
    python tools/publish_release.py --dry-run
"""
from __future__ import annotations

import argparse
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
    return body + "\n"


def is_prerelease(version: str) -> bool:
    u = version.upper()
    return any(tag in u for tag in ("ALPHA", "BETA", "RC", "DEV"))


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


def main() -> int:
    parser = argparse.ArgumentParser(description="Опубликовать GitHub Release")
    parser.add_argument("--dry-run", action="store_true", help="только показать текст, не публиковать")
    args = parser.parse_args()

    version = read_version()
    tag = f"v{version}"
    title = f"OVERTIMETAB {version}"
    notes = changelog_section(version)
    branch = current_branch()
    pre = is_prerelease(version)

    print(f"версия:   {version}")
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
            cmd = [
                "release", "create", tag,
                "--title", title,
                "--notes-file", str(notes_file),
                "--target", branch,
            ]
            if pre:
                cmd.append("--prerelease")
            r = run_gh(cmd, check=False)
            action = "создали"
        if r.returncode != 0:
            sys.stderr.write(r.stderr or r.stdout or "gh не смог опубликовать релиз\n")
            return r.returncode
        print((r.stdout or "").strip() or f"{action} релиз {tag}")
    finally:
        if notes_file.exists():
            notes_file.unlink()

    print("Готово. Zip к релизу приложит Actions, когда дособерёт Windows.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
