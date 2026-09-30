#!/usr/bin/env python3
"""Тест url в version.json: сборка, перезапись и склейка.

1) tools/release_url.py — имя архива из окружения GitHub Actions:
   по тегу — точное имя zip (совпадает с архивом в релизе),
   по ветке — пусто, при неполных данных — пусто.
2) app_update.write_version_json — программа перезаписывает version.json
   при каждом запуске; имя архива из существующего файла обязано
   переноситься в новый (иначе запасной путь обновления теряет архив).
3) app_update.resolve_download_url — имя архива склеивается с адресом
   того хранилища, где лежит version.json:
     GitHub «latest»  → releases/latest/download/<архив>
     GitHub с тегом   → releases/download/<тег>/<архив>
     запасной сервер  → post.mvd.ru/~…/<архив>
   Полный адрес (если вдруг есть) остаётся как есть.
4) app_update._raw_github_version_json — для ссылки «latest» база
   адреса — постоянная ссылка последнего релиза (не битая без тега).

Запуск:
    python3 qa/test_version_json_url.py        # из корня репозитория
"""
import json
import os
import sys
import tempfile
from pathlib import Path

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from release_url import asset_url, from_ci_env  # noqa: E402

import app_update  # noqa: E402

WANT = "OVERTIMETAB_BETA.1_87f3711.zip"


def test_asset_url():
    got = asset_url("BETA.1", "87f3711abc")
    assert got == WANT, got
    assert asset_url("BETA.1", "87f3711") == WANT
    assert asset_url("", "87f3711") == ""
    assert asset_url("BETA.1", "") == ""
    print("asset_url: точное имя архива и пустые случаи ✓")


def test_from_ci_env():
    tag_env = {"GITHUB_REF": "refs/tags/v2.0.0-ALPHA.20",
               "GITHUB_SHA": "87f3711abcde",
               "GITHUB_REPOSITORY": "doc789982-design/OverTimeApp"}
    assert from_ci_env("BETA.1", tag_env) == WANT
    branch_env = dict(tag_env, GITHUB_REF="refs/heads/arena/01a043e7-overtimeapp")
    assert from_ci_env("BETA.1", branch_env) == ""
    print("from_ci_env: по тегу — имя, по ветке — пусто ✓")


def test_preserve_url():
    tmp = Path(tempfile.mkdtemp(prefix="vjson_")) / "version.json"
    shipped = {"name": "OVERTIMETAB", "version": "2.0.0-ALPHA.228",
               "build": 228, "display": "BETA.1", "url": WANT}
    tmp.write_text(json.dumps(shipped, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    # программа при запуске переписывает файл (как Main._init_updates)
    app_update.write_version_json(tmp, "BETA.1", 228)
    after = json.loads(tmp.read_text(encoding="utf-8"))
    assert after.get("url") == WANT, "url потерян при перезаписи: %s" % after
    assert after["build"] == 228 and after["version"] == "BETA.1"
    # явный url имеет приоритет
    other = "OVERTIMETAB_BETA.1_deadbee.zip"
    app_update.write_version_json(tmp, "BETA.1", 229, url=other)
    after2 = json.loads(tmp.read_text(encoding="utf-8"))
    assert after2.get("url") == other, after2
    # нового файла без url — поле не появляется
    tmp2 = tmp.with_name("v2.json")
    app_update.write_version_json(tmp2, "BETA.1", 229)
    assert "url" not in json.loads(tmp2.read_text(encoding="utf-8"))
    print("write_version_json: имя архива переносится/не теряется ✓")


def test_resolve_download_url():
    # запасной сервер (post.mvd.ru): имя склеивается с его адресом
    got = app_update.resolve_download_url(
        "OVERTIMETAB_BETA.1_87f3711.zip",
        "https://post.mvd.ru/~mgrigorev46@mvd.ru/")
    assert got == "https://post.mvd.ru/~mgrigorev46@mvd.ru/OVERTIMETAB_BETA.1_87f3711.zip", got
    # GitHub «latest»: постоянный адрес последнего релиза
    got = app_update.resolve_download_url(
        "OVERTIMETAB_BETA.1_87f3711.zip",
        "https://github.com/doc789982-design/OverTimeApp/releases/latest/download")
    assert got == ("https://github.com/doc789982-design/OverTimeApp/"
                   "releases/latest/download/OVERTIMETAB_BETA.1_87f3711.zip"), got
    # GitHub с тегом
    got = app_update.resolve_download_url(
        "OVERTIMETAB_BETA.1_87f3711.zip",
        "https://github.com/doc789982-design/OverTimeApp/releases/download/v2.0.0-ALPHA.20")
    assert got == ("https://github.com/doc789982-design/OverTimeApp/"
                   "releases/download/v2.0.0-ALPHA.20/OVERTIMETAB_BETA.1_87f3711.zip"), got
    # полный адрес остаётся как есть
    full = "https://example.com/x.zip"
    assert app_update.resolve_download_url(full, "https://other/y") == full
    print("resolve_download_url: имя склеивается с любым хранилищем ✓")


class _FakeResp:
    def __init__(self, raw):
        self._raw = raw

    def read(self):
        return self._raw

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def test_latest_base_url():
    """Запасной GitHub-путь: для «latest» база — постоянный адрес последнего
    релиза (раньше тег подставлялся пустым и ссылка была битой)."""
    saved = app_update.urllib.request.urlopen
    payload = json.dumps({
        "name": "OVERTIMETAB", "version": "2.0.0-ALPHA.231", "build": 231,
        "display": "BETA.1", "url": WANT,
    }).encode("utf-8")

    def fake_urlopen(url, timeout=0):
        return _FakeResp(payload)

    try:
        app_update.urllib.request.urlopen = fake_urlopen
        # ссылка «latest» (как DEFAULT_UPDATE_URL) — тега нет
        info, err = app_update._raw_github_version_json(
            {"owner": "doc789982-design", "repo": "OverTimeApp", "tag": ""})
        assert info, err
        assert info["base_url"] == (
            "https://github.com/doc789982-design/OverTimeApp/releases/latest/download"), info
        assert info["url"] == WANT, info
        # ссылка с тегом
        info2, err2 = app_update._raw_github_version_json(
            {"owner": "doc789982-design", "repo": "OverTimeApp", "tag": "v2.0.0-ALPHA.20"})
        assert info2, err2
        assert info2["base_url"] == (
            "https://github.com/doc789982-design/OverTimeApp/"
            "releases/download/v2.0.0-ALPHA.20"), info2
    finally:
        app_update.urllib.request.urlopen = saved
    print("_raw_github_version_json: latest без тега не даёт битую ссылку ✓")


def main() -> int:
    test_asset_url()
    test_from_ci_env()
    test_preserve_url()
    test_resolve_download_url()
    test_latest_base_url()
    print("═══ URL В version.json: ИМЯ АРХИВА, СКЛЕЙКА И ПЕРЕЗАПИСЬ В ПОРЯДКЕ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
