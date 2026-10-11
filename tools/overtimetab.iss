; -*- coding: utf-8 -*-
; ============================================================
; OVERTIMETAB — стандартный установщик (Inno Setup)
;
; Те самые привычные окна «Далее → Папка → Установить», как у любой
; программы. Сборка на CI (GitHub Actions, windows-latest):
;
;   choco install innosetup -y
;   iscc /F"<имя_файла>" tools\overtimetab.iss
; Версию скрипт берёт из build/product_version.iss — его генерирует
; tools/make_resources.py из AppTheme.qml (шаг «Сгенерировать ресурсы»
; в CI идёт раньше). /DProductVersion с командной строки перекрывается.
;
; /F задаёт имя итогового exe (совпадает с именем zip релиза), сам файл
; кладётся в корень репозитория (OutputDir=..). Версию сборки CI читает
; из components/AppTheme.qml.
;
; Что делает установщик:
;   · спрашивает «для меня / для всех» (при «для всех» — запрос прав);
;   · папка по умолчанию: AppData\Programs\OVERTIMETAB («для меня»)
;     или Program Files\OVERTIMETAB («для всех»);
;   · ярлык в меню «Пуск», по галке — на рабочем столе;
;   · запись в «Установка и удаление программ» и полноценное удаление;
;   · при удалении спрашивает, сносить ли данные сотрудников
;     (Документы\OverTimeTab) — по умолчанию они остаются;
;   · при установке поверх прежней — помнит прежнюю папку.
;
; Обновление установленной программы по-прежнему идёт zip-архивом с
; релиза (автообновление), этот exe — только для людей.
; ============================================================

#ifndef ProductVersion
#define ProductVersion "dev"
#endif
; Честная версия программы — из components/AppTheme.qml (файл генерирует
; tools/make_resources.py при сборке). Подключается ПОСЛЕ аргументов
; командной строки CI и перекрывает их: версия живёт в одном месте.
#include "..\build\product_version.iss"
; папка с собранной программой (относительно этого файла)
#ifndef SourceDir
#define SourceDir "..\dist\OVERTIMETAB"
#endif

; Гарантия версии в релизном zip (баг сборок 294-298: в архиве не было
; номера версии, клиенты не могли обновиться). К этому моменту шаг
; «Собрать ZIP» уже прошёл, а загрузка релиза ещё нет — скрипт
; докладывает version.json в корень архива и сверяет содержимое.
#exec 'python "' + SourcePath + '/stamp_zip_ci.py"

[Setup]
AppId={{7E4C1D52-9A63-4B18-B0F4-6C2D8E5A1F93}
AppName=OVERTIMETAB
AppVersion={#ProductVersion}
AppVerName=OVERTIMETAB {#ProductVersion}
AppPublisher=OVERTIMETAB
DefaultDirName={autopf}\OVERTIMETAB
DefaultGroupName=OVERTIMETAB
; «для меня» по умолчанию; в начале установки спрашивается режим,
; при выборе «для всех» установщик сам перезапустится с правами
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; иконка и вид
SetupIconFile=..\app_icon.ico
UninstallDisplayIcon={app}\OVERTIMETAB.exe
WizardStyle=modern
; сжатие: exe получается заметно меньше суммы частей
Compression=lzma2/max
SolidCompression=yes
; обновление поверх прежней установки помнит её папку
UsePreviousAppDir=yes
; итоговый файл: имя задаётся ключом /F на CI, здесь — для ручной сборки
OutputBaseFilename=OVERTIMETAB_setup
; exe кладём в корень репозитория (папка выше tools\)
OutputDir=..

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
; галка в мастере: ярлык на рабочем столе (в «Пуск» — всегда)
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; \
    GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; \
    Flags: recursesubdirs createallsubdirs ignoreversion
; version.json в корень установленной программы: обновлятор любого
; поколения читает свою версию именно отсюда (спека кладёт копию
; и в _internal — это вторая страховка)
Source: "{#SourcePath}\..\build\version.json"; DestDir: "{app}"; \
    Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\OVERTIMETAB\OVERTIMETAB"; Filename: "{app}\OVERTIMETAB.exe"
Name: "{autodesktop}\OVERTIMETAB"; Filename: "{app}\OVERTIMETAB.exe"; \
    Tasks: desktopicon

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

// При удалении спрашиваем про данные сотрудников. По умолчанию
// (ответ «Нет») базы и отчёты в Документах остаются на месте.
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    if MsgBox(
        'Удалить также данные сотрудников?' + #13#10 +
        'Папка «Документы\OverTimeTab» с базами и отчётами будет удалена.',
        mbConfirmation, MB_YESNO) = IDYES then
    begin
      DelTree(ExpandConstant('{userdocs}\OverTimeTab'), True, True, True);
    end;
  end;
end;
