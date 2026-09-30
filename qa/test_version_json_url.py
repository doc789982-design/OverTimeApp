#!/usr/bin/env python3
"""Тест url в version.json: сборка и перезапись.

1) tools/release_url.py — адрес архива из окружения GitHub Actions:
   по тегу — точная ссылка (совпадает с именем zip в релизе),
   по ветке — пусто, при неполных данных — пусто.
2) app_update.write_version_json — программа перезаписывает version.json
   при каждом запуске; ссылка на архив из существующего файла обязана
   переноситься в новый (иначе запасной путь обновления теряет архив).

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

WANT = ("https://github.com/doc789982-design/OverTimeApp/releases/"
        "download/v2.0.0-ALPHA.20/OVERTIMETAB_BETA.1_87f3711.zip")


def test_asset_url():
    got = asset_url("BETA.1", "87f3711abc", "v2.0.0-ALPHA.20",
                    "doc789982-design/OverTimeApp")
    assert got == WANT, got
    assert asset_url("BETA.1", "87f3711", "refs/tags/v2.0.0-ALPHA.20",
                     "doc789982-design/OverTimeApp") == WANT
    assert asset_url("", "87f3711", "v2", "o/r") == ""
    assert asset_url("BETA.1", "", "v2", "o/r") == ""
    assert asset_url("BETA.1", "87f3711", "", "o/r") == ""
    print("asset_url: точная ссылка и пустые случаи ✓")


def test_from_ci_env():
    tag_env = {"GITHUB_REF": "refs/tags/v2.0.0-ALPHA.20",
               "GITHUB_SHA": "87f3711abcde",
               "GITHUB_REPOSITORY": "doc789982-design/OverTimeApp"}
    assert from_ci_env("BETA.1", tag_env) == WANT
    branch_env = dict(tag_env, GITHUB_REF="refs/heads/arena/01a043e7-overtimeapp")
    assert from_ci_env("BETA.1", branch_env) == ""
    print("from_ci_env: по тегу — ссылка, по ветке — пусто ✓")


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
    other = WANT.replace("87f3711", "deadbee")
    app_update.write_version_json(tmp, "BETA.1", 229, url=other)
    after2 = json.loads(tmp.read_text(encoding="utf-8"))
    assert after2.get("url") == other, after2
    # нового файла без url — поле не появляется
    tmp2 = tmp.with_name("v2.json")
    app_update.write_version_json(tmp2, "BETA.1", 229)
    assert "url" not in json.loads(tmp2.read_text(encoding="utf-8"))
    print("write_version_json: url переносится/не теряется ✓")


def main() -> int:
    test_asset_url()
    test_from_ci_env()
    test_preserve_url()
    print("═══ URL В version.json: СБОРКА И ПЕРЕЗАПИСЬ В ПОРЯДКЕ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
