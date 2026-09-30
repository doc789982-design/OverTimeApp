#!/usr/bin/env python3
"""Имя архива релиза для version.json (внутрь сборки).

Zip на GitHub Actions называется «OVERTIMETAB_<имя>_<sha7>.zip» и
прикрепляется к релизу по тегу. В version.json лежит именно ИМЯ архива,
а не полный адрес: этот файл читают разные хранилища — GitHub и запасной
сервер (post.mvd.ru), — и имя склеивается с адресом того хранилища,
где файл лежит (app_update.resolve_download_url).

Имя появляется только в сборке по тегу (релизной): артефакты сборок
по ветке к релизу не прикрепляются.
"""
import os


def asset_url(display: str, sha: str, tag: str = "", repo: str = "") -> str:
    """Имя zip-архива релиза или пустая строка.

    display — имя версии для человека («BETA.1», из AppTheme),
    sha     — короткий хеш коммита (7 знаков).
    tag и repo не нужны для имени, оставлены для совместимости вызова.
    """
    display = (display or "").strip()
    sha = (sha or "").strip()[:7]
    if not (display and sha):
        return ""
    return "OVERTIMETAB_%s_%s.zip" % (display, sha)


def from_ci_env(display: str, env=None) -> str:
    """Имя архива из окружения GitHub Actions (пустая, если это не релиз)."""
    env = os.environ if env is None else env
    ref = (env.get("GITHUB_REF") or "").strip()
    if not ref.startswith("refs/tags/"):
        return ""                     # сборка по ветке — релизного архива нет
    return asset_url(display, env.get("GITHUB_SHA") or "")
