; -*- coding: utf-8 -*-
; ============================================================
; OVERTIMETAB — стандартный установщик (Inno Setup)
;
; Те самые привычные окна «Далее → Папка → Установить», как у любой
; программы. Сборка на CI (GitHub Actions, windows-latest):
;
;   choco install innosetup -y
;   iscc /DProductVersion="2.0.0-ALPHA.<сборка>" /F"<имя_файла>" tools\overtimetab.iss
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
; папка с собранной программой (относительно этого файла)
#ifndef SourceDir
#define SourceDir "..\dist\OVERTIMETAB"
#endif

[Setup]
AppId={{7E4C1D52-9A63-4B18-B0F4-6C2D8E5A1F93}
AppName=OVERTIMETAB
AppVersion={#ProductVersion}
AppVerName=OVERTIMETAB (сборка {#ProductVersion})
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
