#!/usr/bin/env python3
"""Тест схемы версий: имя для человека + семвер + непрерывные сборки.

В AppTheme.qml живут два поля: appVersion — имя для человека (до релиза
просто «BETA.1», без номера версии: версия появится на релизе, дальше —
нумерация; его же берут имена файлов релиза и установщик) и appSemver —
семвер для механики (тег v<семвер>, сравнение версий, машинная строка).
Сборки непрерывные, никогда не сбрасываются; «новее ли» обновление решает
сборка, а у самых старых клиентов — машинная строка, где счётчик канала =
номер сборки.

Проверяется:
  A. Парсер: каналы по старшинству alpha < beta < rc < релиз,
     релиз старше пре-релиза, ядро 2.0.0 < 2.1.0, счётчики числом;
  B. Машинная строка (scan_version): счётчик канала = сборка,
     релизная версия не меняется;
  C. Обновление со старых сборок: клиент с номером сборки видит
     новую версию любой канальности; строковый клиент ALPHA видит
     beta/релиз/2.1.0 новее и не видит отката;
  D. Согласованность проекта: appVersion — имя (не семвер), appSemver —
     семвер; updates/version.json совпадает с AppTheme (машинная строка
     из семвера, display = имя, build); spec читает оба поля; установщик
     получает имя через сгенерированный include; тег релиза = v + семвер,
     заголовок — из имени; в нигде нет зашитой версии;
  E. Зип-сканер видит машинную строку (и не видит человеческие имена);
  F. Ключ чейнджлога «ВЕРСИЯ+сборка» читается (включая старый формат).

Запуск:
    python3 qa/test_version_scheme.py
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import app_update  # noqa: E402
from release_url import asset_url  # noqa: E402

SEMVER_RE = re.compile(r"^\d+\.\d+\.\d+(-(alpha|beta|rc)\.\d+)?$")


def read_identity() -> tuple[str, str, int]:
    theme = (ROOT / "components" / "AppTheme.qml").read_text(encoding="utf-8")
    m = re.search(r'appVersion:\s*"([^"]+)"', theme)
    s = re.search(r'appSemver:\s*"([^"]+)"', theme)
    b = re.search(r"appBuild:\s*(\d+)", theme)
    assert m and s and b, "в AppTheme.qml нет appVersion/appSemver/appBuild"
    return m.group(1), s.group(1), int(b.group(1))


def main() -> int:
    name, semver, build = read_identity()

    # ── A. Парсер ──
    p = app_update.parse_version
    assert p("2.0.0") == ((2, 0, 0), "", 0)
    assert p("2.0.0-alpha.3") == ((2, 0, 0), "ALPHA", 3)
    assert p("2.0.0-beta.1") == ((2, 0, 0), "BETA", 1)
    assert p("2.0.0-rc.2") == ((2, 0, 0), "RC", 2)
    assert p("2.1.0-alpha.1") == ((2, 1, 0), "ALPHA", 1)
    # старшинство: релиз > rc > beta > alpha одного ядра
    assert not app_update.is_newer("2.0.0-rc.1", "2.0.0")
    assert app_update.is_newer("2.0.0", "2.0.0-rc.1")
    assert app_update.is_newer("2.0.0-rc.1", "2.0.0-beta.9")
    assert app_update.is_newer("2.0.0-beta.2", "2.0.0-alpha.7")
    # ядро старше любого канала соседнего ядра
    assert app_update.is_newer("2.1.0-alpha.1", "2.0.0")
    assert not app_update.is_newer("2.0.0", "2.1.0-alpha.1")
    # счётчики сравниваются числом, не строкой
    assert app_update.is_newer("2.0.0-beta.10", "2.0.0-beta.9")
    assert not app_update.is_newer("2.0.0-beta.2", "2.0.0-beta.10")
    print("A: парсер — каналы alpha < beta < rc < релиз, ядро, счётчики ✓")

    # ── B. Машинная строка ──
    scan = app_update.scan_version
    assert scan("2.0.0-beta.1", 294) == "2.0.0-beta.294"
    assert scan("2.0.0-alpha.5", 70) == "2.0.0-alpha.70"
    assert scan("2.0.0", 300) == "2.0.0", "релиз не меняется"
    assert scan("2.1.0-alpha.1", 310) == "2.1.0-alpha.310"
    assert scan("", 5) == "" and scan("2.0.0-beta.1", 0) == "2.0.0-beta.1"
    print("B: машинная строка — счётчик канала = сборка, релиз как есть ✓")

    # ── C. Обновление со старых сборок ──
    # клиент со сборкой (все недавние): решает сборка, канал не важен
    assert app_update.is_newer("2.0.0-beta.1", "BETA.1", 295, 294)
    assert app_update.is_newer("2.0.0", "BETA.1", 300, 295)
    assert app_update.is_newer("2.1.0-alpha.1", "2.0.0", 310, 300)
    assert not app_update.is_newer("2.0.0-beta.2", "BETA.1", 295, 295)
    assert not app_update.is_newer("2.0.0-beta.2", "BETA.1", 294, 295)
    # строковый клиент (самые старые, сборки не знают): машинная строка
    # против их «2.0.0-ALPHA.N» — канал растёт, счётчик = сборка
    assert app_update.is_newer("2.0.0-beta.294", "2.0.0-ALPHA.293")
    assert app_update.is_newer("2.0.0", "2.0.0-ALPHA.293")
    assert app_update.is_newer("2.1.0-alpha.310", "2.0.0")
    assert not app_update.is_newer("2.0.0-beta.294", "2.0.0-beta.294")
    assert not app_update.is_newer("2.0.0-alpha.5", "2.0.0-beta.2")
    print("C: старые клиенты видят новую версию, откат запрещён ✓")

    # ── D. Согласованность проекта ──
    # имя — не семвер (просто «BETA.1»), семвер — семвер
    assert name and not SEMVER_RE.match(name), \
        f"appVersion «{name}» — это семвер, а должно быть имя «BETA.1» " \
        "(без номера версии — версия появится на релизе)"
    assert SEMVER_RE.match(semver), \
        f"appSemver «{semver}» — не семвер с каналом (см. RELEASING.md)"
    assert build > 0
    vj = json.loads((ROOT / "updates" / "version.json")
                    .read_text(encoding="utf-8"))
    assert vj.get("build") == build, \
        f"build в version.json ({vj.get('build')}) != appBuild ({build})"
    assert vj.get("display") == name, \
        f"display в version.json ({vj.get('display')}) != appVersion"
    assert vj.get("version") == app_update.scan_version(semver, build), \
        f"version в version.json ({vj.get('version')}) — не машинная строка"
    url = str(vj.get("url") or "")
    assert re.fullmatch(r"OVERTIMETAB_[^\s/]+\.zip", url), \
        "url в version.json — не имя архива (правило: имя, не адрес)"
    assert url.startswith(f"OVERTIMETAB_{name}_"), \
        f"имя архива строится из appVersion: {url}"
    # spec читает оба поля и вычисляет машинную строку из семвера
    spec = (ROOT / "tools" / "overtimetab.spec").read_text(encoding="utf-8")
    assert "appSemver" in spec, "spec не читает appSemver"
    assert "_app_semver" in spec and "scan_version(_app_semver" in spec, \
        "машинная версия в spec считается не из семвера"
    assert "2.0.0-ALPHA" not in spec, "в spec осталась зашитая версия"
    # установщик получает имя через сгенерированный include:
    # make_resources.py пишет build/product_version.iss из appVersion,
    # а overtimetab.iss подключает его ПОСЛЕ аргументов командной строки
    # CI — версия из /D перекрывается (workflow менять из песочницы нельзя)
    mr = (ROOT / "tools" / "make_resources.py").read_text(encoding="utf-8")
    assert "product_version.iss" in mr and "appVersion" in mr, \
        "make_resources не доводит имя до установщика"
    iss = (ROOT / "tools" / "overtimetab.iss").read_text(encoding="utf-8")
    assert '#include "..\\build\\product_version.iss"' in iss, \
        "установщик не подключает сгенерированную версию"
    assert "#undef ProductVersion" in mr, \
        "сгенерированный файл не перекрывает аргументы командной строки"
    assert "2.0.0-ALPHA" not in iss, "в установщике осталась зашитая версия"
    # публикатор: тег из семвера, заголовок из имени
    pub = (ROOT / "tools" / "publish_release.py").read_text(encoding="utf-8")
    assert "def read_semver" in pub, "публикатор не читает семвер"
    assert 'f"v{semver}"' in pub, "тег релиза строится не из семвера"
    assert 'title = f"OVERTIMETAB {version}"' in pub, \
        "заголовок релиза строится не из имени"
    # имя архива строится из имени
    assert asset_url(name, "87f3711") == f"OVERTIMETAB_{name}_87f3711.zip"
    # RELEASING объясняет схему: два поля, каналы, сборки, тег
    doc = (ROOT / "RELEASING.md").read_text(encoding="utf-8")
    for word in ("Схема версий", "appVersion", "appSemver", "alpha", "beta",
                 "сборк", "v` + семвер"):
        assert word in doc, "RELEASING.md не объясняет: " + word
    print(f"D: проект согласован — имя {name}, семвер {semver}, "
          f"сборка {build}, нигде нет зашитой версии ✓")

    # ── E. Зип-сканер ──
    looks = app_update._looks_like_app_version
    assert looks(app_update.scan_version(semver, build)), \
        "сканер зипов не видит машинную строку"
    assert looks("2.0.0") and not looks("BETA.1"), \
        "сканер зипов понимает только X.Y.Z-версии"
    print("E: сканер зипов видит машинную строку ✓")

    # ── F. Ключ чейнджлога ──
    key = app_update.build_from_version_key
    assert key(f"{name}+{build}") == build
    assert key("BETA.1+239") == 239, "старый формат ключа должен читаться"
    assert key("2.0.0-beta.1") == 0
    print("F: ключ «ВЕРСИЯ+сборка» читается (и старый формат тоже) ✓")

    print("═══ СХЕМА ВЕРСИЙ: ИМЯ + СЕМВЕР + НЕПРЕРЫВНЫЕ СБОРКИ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
