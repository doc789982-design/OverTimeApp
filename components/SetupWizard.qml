import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// МАСТЕР УСТАНОВКИ / УДАЛЕНИЯ OVERTIMETAB
//
// Запускается заглушкой релизного exe (распаковывает программу в
// временную папку и зовёт OVERTIMETAB.exe --setup) либо из
// «Установка и удаление программ» (--uninstall). Логика — installer.py,
// мост в QML — SetupBackend в Main.py.
//
// Страницы: приветствие → (найдена копия?) → место и ярлыки →
// прогресс → готово. Удаление: подтверждение → готово.
// ============================================================

Item {
    id: wizard
    // Окно (размеры, заголовок, показ) задаёт main_setup.qml — компонент
    // из каталога-модуля components сам себе окном быть не может: контекст
    // QML даёт только входной файл из корня (как main.qml программы).

    // ── состояние ──
    property string page: "welcome"   // welcome | place | progress | done
    property bool perUser: true       // «для меня» / «для всех»
    property bool foundShown: false   // уже показали страницу с находкой
    property string installDir: ""
    property bool makeDesktop: true
    property bool makeStartMenu: true
    property bool removeData: false
    property bool runAfter: true
    property string errorText: ""

    function safeDir(s): string {
        s = (s || "").trim()
        if (!s) return ""
        s = s.replace(/\//g, "\\")
        while (s.endsWith("\\")) s = s.slice(0, -1)
        return s
    }

    function startInstall() {
        var d = safeDir(dirField.text)
        if (!d || d.length < 3 || d === "C:\\" || d === "c:\\") {
            errorText = "Укажите папку установки — например, " + backend.defaultDir
            return
        }
        errorText = ""
        installDir = d
        page = "progress"
        backend.install(d, perUser, makeDesktop, makeStartMenu)
    }

    Component.onCompleted: {
        installDir = backend.defaultDir
        dirField.text = installDir
        if (backend.mode === "uninstall")
            page = "confirmRemove"
    }

    // фон — единая поверхность, без теней (правило тёмной темы)
    Rectangle { anchors.fill: parent; color: AppTheme.bgPanel }

    // сигналы из потока установки
    Connections {
        target: backend
        function onFinishedOk() {
            wizard.page = "done"
        }
        function onFinishedError(msg) {
            wizard.errorText = msg
        }
        function onRequestQuit() {
            Qt.quit()
        }
    }

    ColumnLayout {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceL
        spacing: AppTheme.spaceM

        // ── шапка ──
        ColumnLayout {
            Layout.fillWidth: true
            spacing: 4
            Label {
                text: backend.mode === "uninstall"
                      ? "Удаление OVERTIMETAB"
                      : (wizard.page === "done" ? "Установка завершена"
                         : "Установка OVERTIMETAB")
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: 22
                font.weight: Font.DemiBold
            }
            Label {
                text: backend.mode === "uninstall"
                      ? backend.existingBuild > 0
                        ? "Сборка " + backend.existingBuild + " · " + backend.existingDir
                        : backend.existingDir
                      : AppTheme.appVersionFull
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: 13
                elide: Text.ElideMiddle
                Layout.fillWidth: true
            }
        }

        // ── контент ──
        StackLayout {
            id: stack
            Layout.fillWidth: true
            Layout.fillHeight: true
            currentIndex: page === "welcome" ? 0
                          : page === "place" ? 1
                          : page === "progress" ? 2
                          : page === "confirmRemove" ? 3 : 4

            // 0 · приветствие
            ColumnLayout {
                spacing: AppTheme.spaceM
                Label {
                    Layout.fillWidth: true
                    wrapMode: Text.WordWrap
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 15
                    text: "Программа учёта служебного времени. Базы и настройки "
                          + "хранятся отдельно (в «Документах») — при обновлении "
                          + "ничего не теряется."
                }
                // найденная копия — карточкой
                Rectangle {
                    visible: backend.existingFound && !foundShown
                    Layout.fillWidth: true
                    implicitHeight: col.height + AppTheme.spaceM * 2
                    radius: AppTheme.radiusMedium
                    color: AppTheme.bgSurface
                    border.color: AppTheme.bgCell
                    ColumnLayout {
                        id: col
                        anchors.fill: parent
                        anchors.margins: AppTheme.spaceM
                        spacing: 6
                        Label {
                            text: "На этом компьютере уже установлена копия" +
                                  (backend.existingBuild > 0
                                   ? " — сборка " + backend.existingBuild : "")
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: 14
                            font.weight: Font.DemiBold
                            wrapMode: Text.WordWrap
                            Layout.fillWidth: true
                        }
                        Label {
                            text: backend.existingDir
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: 13
                            elide: Text.ElideMiddle
                            Layout.fillWidth: true
                        }
                    }
                }
                Item { Layout.fillHeight: true }
            }

            // 1 · место установки и ярлыки
            ColumnLayout {
                spacing: AppTheme.spaceM
                Label {
                    text: "Куда установить"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 15
                    font.weight: Font.DemiBold
                }
                // режим размещения
                ColumnLayout {
                    spacing: 2
                    AppRadioButton {
                        text: "Для меня (рекомендуется) — без прав администратора"
                        checked: wizard.perUser
                        onToggled: { wizard.perUser = true; dirField.text = backend.dirFor(true) }
                    }
                    AppRadioButton {
                        text: "Для всех пользователей (папка Program Files)"
                        checked: !wizard.perUser
                        onToggled: { wizard.perUser = false; dirField.text = backend.dirFor(false) }
                    }
                }
                // папка
                ColumnLayout {
                    spacing: 4
                    Layout.fillWidth: true
                    Label {
                        text: "Папка установки"
                        color: AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: 12
                    }
                    Rectangle {
                        Layout.fillWidth: true
                        height: 40
                        radius: AppTheme.radiusSmall
                        color: AppTheme.bgSurface
                        border.color: errorText ? AppTheme.accentDanger : AppTheme.bgCell
                        TextInput {
                            id: dirField
                            anchors.fill: parent
                            anchors.margins: 10
                            verticalAlignment: TextInput.AlignVCenter
                            color: AppTheme.textPrimary
                            selectionColor: AppTheme.accentBrand
                            selectedTextColor: AppTheme.textOnAccent
                            font.family: AppTheme.fontFamily
                            font.pixelSize: 13
                            clip: true
                            text: wizard.installDir
                            onAccepted: wizard.startInstall()
                        }
                    }
                }
                // ярлыки
                Label {
                    text: "Ярлыки"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 15
                    font.weight: Font.DemiBold
                }
                ColumnLayout {
                    spacing: 2
                    AppCheckBox {
                        text: "На рабочем столе"
                        checked: wizard.makeDesktop
                        onToggled: wizard.makeDesktop = checked
                    }
                    AppCheckBox {
                        text: "В меню «Пуск»"
                        checked: wizard.makeStartMenu
                        onToggled: wizard.makeStartMenu = checked
                    }
                }
                Label {
                    visible: wizard.errorText
                    text: wizard.errorText
                    color: AppTheme.accentDanger
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 12
                    wrapMode: Text.WordWrap
                    Layout.fillWidth: true
                }
                Item { Layout.fillHeight: true }
            }

            // 2 · прогресс
            ColumnLayout {
                spacing: AppTheme.spaceM
                Item { Layout.fillHeight: true }
                Label {
                    Layout.alignment: Qt.AlignHCenter
                    text: backend.mode === "uninstall" ? "Удаляем…" : "Устанавливаем…"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 16
                    font.weight: Font.DemiBold
                }
                ProgressBar {
                    id: bar
                    Layout.fillWidth: true
                    from: 0; to: 1; value: backend.progress
                    background: Rectangle {
                        implicitHeight: 6
                        radius: 3
                        color: AppTheme.bgCell
                    }
                    contentItem: Item {
                        implicitHeight: 6
                        Rectangle {
                            width: Math.max(6, bar.visualPosition * parent.width)
                            height: parent.height
                            radius: 3
                            color: AppTheme.accentBrand
                        }
                    }
                }
                Label {
                    Layout.alignment: Qt.AlignHCenter
                    text: backend.statusLine
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 12
                }
                Label {
                    visible: wizard.errorText
                    Layout.fillWidth: true
                    wrapMode: Text.WordWrap
                    horizontalAlignment: Text.AlignHCenter
                    text: wizard.errorText
                    color: AppTheme.accentDanger
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 12
                }
                Item { Layout.fillHeight: true }
            }

            // 3 · подтверждение удаления
            ColumnLayout {
                spacing: AppTheme.spaceM
                Label {
                    Layout.fillWidth: true
                    wrapMode: Text.WordWrap
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 15
                    text: "Программа будет удалена с компьютера."
                }
                AppCheckBox {
                    text: "Удалить также данные сотрудников — «Документы\\OverTimeTab»"
                    checked: wizard.removeData
                    onToggled: wizard.removeData = checked
                }
                Label {
                    Layout.fillWidth: true
                    wrapMode: Text.WordWrap
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 12
                    text: wizard.removeData
                          ? "Вместе с базами и отчётами — отменить будет нельзя."
                          : "Базы и отчёты останутся в «Документах\\OverTimeTab»."
                }
                Item { Layout.fillHeight: true }
            }

            // 4 · готово
            ColumnLayout {
                spacing: AppTheme.spaceM
                Item { Layout.fillHeight: true }
                Label {
                    Layout.alignment: Qt.AlignHCenter
                    text: backend.mode === "uninstall" ? "Программа удалена" : "Готово"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 18
                    font.weight: Font.DemiBold
                }
                Label {
                    Layout.alignment: Qt.AlignHCenter
                    visible: backend.mode !== "uninstall"
                    wrapMode: Text.WordWrap
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: 13
                    text: wizard.installDir
                }
                AppCheckBox {
                    visible: backend.mode !== "uninstall"
                    Layout.alignment: Qt.AlignHCenter
                    text: "Запустить программу"
                    checked: wizard.runAfter
                    onToggled: wizard.runAfter = checked
                }
                Item { Layout.fillHeight: true }
            }
        }

        // ── кнопки ──
        RowLayout {
            Layout.fillWidth: true
            spacing: AppTheme.spaceS
            Item { Layout.fillWidth: true }

            AppButton {
                visible: wizard.page === "welcome"
                text: "Отмена"
                variant: "ghost"
                onClicked: Qt.quit()
            }
            AppButton {
                // найдена копия — можно отказаться и выбрать место руками
                visible: wizard.page === "welcome"
                         && backend.existingFound && !foundShown
                text: "Выбрать другое место"
                variant: "ghost"
                onClicked: {
                    wizard.foundShown = true
                    wizard.page = "place"
                    dirField.text = backend.dirFor(wizard.perUser)
                }
            }
            AppButton {
                visible: wizard.page === "welcome"
                text: backend.existingFound && !foundShown
                      ? "Обновить там же" : "Далее"
                onClicked: {
                    if (backend.existingFound && !foundShown) {
                        // обновление найденной копии на месте
                        wizard.perUser = backend.existingPerUser
                        dirField.text = backend.existingDir
                        wizard.page = "progress"
                        backend.installExisting()
                    } else {
                        wizard.foundShown = true
                        wizard.page = "place"
                        dirField.text = backend.dirFor(wizard.perUser)
                    }
                }
            }
            AppButton {
                visible: wizard.page === "place"
                text: "Назад"
                variant: "ghost"
                onClicked: wizard.page = "welcome"
            }
            AppButton {
                visible: wizard.page === "place"
                text: "Установить"
                onClicked: wizard.startInstall()
            }
            AppButton {
                visible: wizard.page === "progress" && wizard.errorText
                text: "Назад"
                variant: "ghost"
                onClicked: { wizard.errorText = ""; wizard.page = "place" }
            }
            AppButton {
                visible: wizard.page === "confirmRemove"
                text: "Отмена"
                variant: "ghost"
                onClicked: Qt.quit()
            }
            AppButton {
                visible: wizard.page === "confirmRemove"
                text: "Удалить"
                variant: "danger"
                onClicked: { wizard.page = "progress"; backend.uninstall(wizard.removeData) }
            }
            AppButton {
                visible: wizard.page === "done"
                text: backend.mode === "uninstall"
                      ? "Закрыть" : (wizard.runAfter ? "Запустить" : "Закрыть")
                onClicked: backend.finish(wizard.runAfter)
            }
        }
    }
}
