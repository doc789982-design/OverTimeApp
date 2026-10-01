# Один шаг руками: установщик Inno Setup в workflow

GitHub не позволяет токену агента менять файлы в `.github/workflows/`
(только владелец через веб). Ниже — единственная нужная правка, после
которой релиз собирается со стандартным установщиком Inno Setup.

## Что заменить

Откройте
[.github/workflows/build-windows.yml](https://github.com/doc789982-design/OverTimeApp/edit/main/.github/workflows/build-windows.yml)
и найдите два шага (стоят между «Собрать ZIP» и «Загрузить сборку в
Artifacts»):

```yaml
      - name: Собрать заглушку установщика
        run: pyinstaller --noconfirm --clean --distpath build_stub tools/installer_stub.spec

      - name: Склеить установщик (заглушка + архив)
        run: python tools/make_installer.py --stub build_stub/installer_stub.exe --zip "${{ steps.ver.outputs.zipname }}.zip" --out "${{ steps.ver.outputs.zipname }}.exe"
```

(между ними может быть комментарий «── Установщик ──…» — удалите и его).

Замените их на один шаг:

```yaml
      - name: Собрать установщик (Inno Setup)
        shell: pwsh
        run: |
          choco install innosetup -y --no-progress
          $theme = Get-Content components/AppTheme.qml -Raw -Encoding UTF8
          $build = if ($theme -match 'appBuild:\s*(\d+)') { $Matches[1] } else { "0" }
          iscc /DProductVersion="2.0.0-ALPHA.$build" /F"${{ steps.ver.outputs.zipname }}" tools\overtimetab.iss
```

Больше ничего не трогайте: exe ложится в корень репозитория под тем же
именем, что и раньше, поэтому шаги Artifacts и релиза не меняются.

## Затем скопировать в рабочую ветку

1. Откройте файл на main (уже с правкой):
   https://github.com/doc789982-design/OverTimeApp/blob/main/.github/workflows/build-windows.yml
2. Кнопка **Copy raw content** (иконка копирования над содержимым).
3. Откройте редактор рабочей ветки:
   https://github.com/doc789982-design/OverTimeApp/edit/arena/01a043e7-overtimeapp/.github/workflows/build-windows.yml
4. Клик в текст → **Ctrl+A, Ctrl+V** → **Commit changes**.

После этого скажите агенту «готово» — он прогонит ворота и опубликует
релиз сборки 244 с обоими файлами. Точный git-патч — рядом:
`tools/inno_workflow.patch`.

## Проверка руками (не обязательно)

После публикации на релизе появится `OVERTIMETAB_BETA.1_<sha>.exe` —
скачайте и запустите: мастер должен выглядеть как установка обычной
Windows-программы (окно «Добро пожаловать», выбор папки, галка ярлыка,
прогресс, «Завершение»). Запись появится в «Установка и удаление
программ».
