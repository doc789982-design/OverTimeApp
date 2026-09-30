#!/usr/bin/env python3
"""Адрес архива релиза для version.json внутри сборки.

Zip на GitHub Actions называется «OVERTIMETAB_<имя>_<sha7>.zip» и
прикрепляется к релизу по тегу. Чтобы version.json, лежащий ВНУТРИ
архива, содержал полную ссылку на самого себя, адрес собирается из
переменных окружения Actions (доступны и pyinstaller-у, который
пишет этот version.json по нашему spec-файлу).

Ссылка появляется только в сборке по тегу (релизной): артефакты
сборок по ветке адреса на GitHub не имеют.
"""
import os


def asset_url(display: str, sha: str, tag: str, repo: str) -> str:
    """Полная ссылка на zip-архив релиза или пустая строка.

    display — имя версии для человека («BETA.1», из AppTheme),
    sha     — короткий хеш коммита (7 знаков),
    tag     — тег релиза («v2.0.0-ALPHA.20»),
    repo    — «owner/name».
    """
    display = (display or "").strip()
    sha = (sha or "").strip()[:7]
    tag = (tag or "").strip().replace("refs/tags/", "")
    repo = (repo or "").strip().strip("/")
    if not (display and sha and tag and repo):
        return ""
    return "https://github.com/%s/releases/download/%s/OVERTIMETAB_%s_%s.zip" % (
        repo, tag, display, sha)


def from_ci_env(display: str, env=None) -> str:
    """Ссылка из окружения GitHub Actions (пустая, если это не релиз)."""
    env = os.environ if env is None else env
    ref = (env.get("GITHUB_REF") or "").strip()
    if not ref.startswith("refs/tags/"):
        return ""                     # сборка по ветке — релизного адреса нет
    return asset_url(
        display,
        env.get("GITHUB_SHA") or "",
        ref,
        env.get("GITHUB_REPOSITORY") or "",
    )
