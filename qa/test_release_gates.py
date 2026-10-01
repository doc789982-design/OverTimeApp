#!/usr/bin/env python3
"""Релизные ворота: релиз без установщика невозможно выпустить.

Проверяется «насильная» часть схемы (запрос владельца проекта: сборка
должна собираться ТОЛЬКО вместе с установщиком, чтобы любой, кто
продолжит сопровождение — человек или ИИ без памяти о договорённостях,
— не смог сделать релиз по-старому, одним зипом):

  A. workflow «Сборка Windows» собирает заглушку, клеит exe и грузит
     в релиз ОБА файла (zip + exe);
  B. publish_release.py проверяет состав релиза: asset_errors ругается
     на отсутствие zip или exe и на чужой sha, принимает полный набор;
  C. publish_release.py дожидается сборки и чистит старые ассеты
     (по маркерам в коде — wait_for_build, cleanup_old_release_assets,
     вызовы в main);
  D. RELEASING.md существует и объясняет схему «zip + exe»;
  E. все части цепочки на месте: заглушка, её spec, склейщик, логика
     установки, окно мастера, вход --setup в Main.py.

Запуск:
    python3 qa/test_release_gates.py
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from publish_release import asset_errors  # noqa: E402


def main() -> int:
    wf = (ROOT / ".github" / "workflows" / "build-windows.yml").read_text(
        encoding="utf-8")
    pub = (ROOT / "tools" / "publish_release.py").read_text(encoding="utf-8")

    # ── A. workflow собирает оба файла ──
    for marker in ("installer_stub.spec",
                   "make_installer.py",
                   "build_stub/installer_stub.exe"):
        assert marker in wf, "в workflow нет шага: " + marker
    up = wf[wf.find("actions/upload-artifact"):]
    assert ".zip" in up and ".exe" in up, "в artifacts нет пары zip+exe"
    rel = wf[wf.find("softprops/action-gh-release"):]
    assert ".zip" in rel and ".exe" in rel, "в релиз грузят не оба файла"
    print("A: workflow собирает заглушку, клеит exe, грузит zip + exe ✓")

    # ── B. проверка состава релиза ──
    ok = [{"name": "OVERTIMETAB_BETA.1_a1b2c3d.zip"},
          {"name": "OVERTIMETAB_BETA.1_a1b2c3d.exe"}]
    assert asset_errors(ok, "a1b2c3d") == [], asset_errors(ok, "a1b2c3d")
    for broken, why in (
            ([{"name": "OVERTIMETAB_BETA.1_a1b2c3d.zip"}], "нет установщика"),
            ([{"name": "OVERTIMETAB_BETA.1_a1b2c3d.exe"}], "нет архива"),
            ([{"name": "OVERTIMETAB_BETA.1_a1b2c3d.zip"},
              {"name": "OVERTIMETAB_BETA.1_fffffff.exe"}], "чужой sha"),
            ([], "пусто")):
        errs = asset_errors(broken, "a1b2c3d")
        assert errs, "проверка должна была ругаться: " + why
        if why == "нет установщика":
            assert any(".exe" in e for e in errs)
        if why == "нет архива":
            assert any(".zip" in e for e in errs)
    print("B: asset_errors ловит отсутствие zip/exe и чужой sha ✓")

    # ── C. публикатор ждёт и чистит ──
    for marker in ("def wait_for_build", "def cleanup_old_release_assets",
                   "asset_errors(assets", "wait_for_build()"):
        assert marker in pub, "в publish_release.py нет: " + marker
    assert "cleanup_old_release_zips" not in pub, \
        "осталась старая чистка только зипов"
    print("C: publish_release дожидается сборки и чистит старые ассеты ✓")

    # ── D. инструкция ──
    doc = (ROOT / "RELEASING.md").read_text(encoding="utf-8")
    for word in (".exe", ".zip", "make_installer", "publish_release"):
        assert word in doc, "RELEASING.md не объясняет: " + word
    print("D: RELEASING.md описывает схему «zip + exe» ✓")

    # ── E. все части цепочки ──
    for p in ("tools/installer_stub.py", "tools/installer_stub.spec",
              "tools/make_installer.py", "installer.py",
              "components/SetupWizard.qml"):
        assert (ROOT / p).exists(), "нет файла " + p
    main_src = (ROOT / "Main.py").read_text(encoding="utf-8")
    assert '"--setup" in sys.argv' in main_src and "setup_main" in main_src
    assert "--uninstall" in main_src
    spec = (ROOT / "tools" / "overtimetab.spec").read_text(encoding="utf-8")
    assert "BUILD_HISTORY" in spec  # карта сборок — в каждой сборке
    print("E: заглушка, склейщик, логика, мастер и вход --setup на месте ✓")

    print("═══ РЕЛИЗ БЕЗ УСТАНОВЩИКА НЕ ВЫПУСТИТЬ: ВОРОТА ДЕРЖАТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
