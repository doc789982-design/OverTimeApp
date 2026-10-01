# Один шаг руками владельца: шаги установщика в workflow

GitHub не позволяет токену агента менять файлы в `.github/workflows/`
(только владелец репозитория через веб). Поэтому правку ниже нужно
один раз внести руками: открыть
[.github/workflows/build-windows.yml](https://github.com/doc789982-design/OverTimeApp/edit/main/.github/workflows/build-windows.yml)
и вставить между шагами «Собрать ZIP» и «Загрузить сборку в Artifacts»:

```yaml
      # ── Установщик ──────────────────────────────────────────
      # Релиз = zip (обновление для программ) + exe (установщик).
      - name: Собрать заглушку установщика
        run: pyinstaller --noconfirm --clean --distpath build_stub tools/installer_stub.spec

      - name: Склеить установщик (заглушка + архив)
        run: python tools/make_installer.py --stub build_stub/installer_stub.exe --zip "${{ steps.ver.outputs.zipname }}.zip" --out "${{ steps.ver.outputs.zipname }}.exe"
```

И заменить в двух местах списки файлов — в «Загрузить сборку в
Artifacts» (path) и «Опубликовать релиз» (files) — на:

```yaml
          files: |
            ${{ steps.ver.outputs.zipname }}.zip
            ${{ steps.ver.outputs.zipname }}.exe
```

(для шага upload-artifact поле называется `path:`, для релиза — `files:`;
многострочное значение `|` обязательно).

После сохранения: скажите агенту — он подтянет правку, пересоберёт тег
и опубликует релиз с обоими файлами. Пока правки нет, тест
`qa/test_release_gates.py` красный — это сознательные ворота: релиз без
установщика не выпускается. Точный git-патч лежит рядом:
`tools/installer_workflow.patch`.
