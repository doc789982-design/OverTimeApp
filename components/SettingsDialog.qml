import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import QtQuick.Dialogs
import Qt5Compat.GraphicalEffects

AppLargeModal {
    id: root
    title: "Настройки программы"

    signal requestFileAttach()

    // ==========================================
    // ПРОПОРЦИОНАЛЬНЫЙ РАЗМЕР
    // Окно настроек масштабируется относительно главного окна программы:
    // на маленьком окне настроек меньше, на развёрнутом — больше, но не
    // бесконечно (ограничены минимумом/максимумом, чтобы не было неудобно).
    // ==========================================
    readonly property int minW: 760
    readonly property int maxW: 1200
    readonly property int minH: 520
    readonly property int maxH: 880

    width:  ApplicationWindow.window
        ? Math.max(minW, Math.min(maxW, ApplicationWindow.window.width  * 0.88))
        : maxW
    height: ApplicationWindow.window
        ? Math.max(minH, Math.min(maxH, ApplicationWindow.window.height * 0.88))
        : maxH

    function openTab(index) {
        settingsMenuList.currentIndex = index
        settingsStack.currentIndex = index
    }

    Row {
        anchors.fill: parent

        // ==========================================
        // ЛЕВОЕ МЕНЮ
        // ==========================================
        Rectangle {
            id: settingsMenuPanel
            width: Math.min(240, parent.width * 0.24)
            height: parent.height
            color: AppTheme.bgPanel

            Rectangle {
                width: 1
                color: AppTheme.borderDivider
                anchors.right: parent.right
                anchors.top: parent.top
                anchors.bottom: parent.bottom
            }

            Column {
                anchors.fill: parent
                anchors.topMargin: AppTheme.spaceL
                anchors.bottomMargin: AppTheme.spaceL
                spacing: 0

                Item {
                    width: parent.width
                    height: 48

                    Text {
                        anchors.left: parent.left
                        anchors.leftMargin: AppTheme.spaceL
                        anchors.verticalCenter: parent.verticalCenter
                        text: "РАЗДЕЛЫ"
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeMicro
                        font.weight: AppTheme.weightBold
                        font.letterSpacing: 1.2
                    }
                }

                ListView {
                    id: settingsMenuList
                    width: parent.width
                    height: contentHeight
                    interactive: false
                    property int currentIndex: 0

                    model: [
                        { title: "Внешний вид",          icon: "../icons/layout.svg",    desc: "Тема и элементы" },
                        { title: "Горячие клавиши",      icon: "../icons/command.svg",   desc: "Сочетания клавиш" },
                        { title: "Управление базами",    icon: "../icons/database.svg",  desc: "Подключение и экспорт" },
                        { title: "Данные подразделения", icon: "../icons/briefcase.svg", desc: "Реквизиты отдела" },
                        { title: "Уведомления",          icon: "../icons/bell.svg",      desc: "Напоминания программы" },
                        { title: "Обновление",           icon: "../icons/sparkle.svg",   desc: "Новая версия с флешки" }
                    ]

                    delegate: Item {
                        width: ListView.view.width
                        height: 56
                        property bool isSelected: settingsMenuList.currentIndex === index

                        Rectangle {
                            visible: isSelected
                            anchors.left: parent.left
                            anchors.verticalCenter: parent.verticalCenter
                            width: 3
                            height: 32
                            radius: 2
                            color: AppTheme.accentBrand
                        }

                        Rectangle {
                            id: itemBg
                            anchors.fill: parent
                            anchors.leftMargin: 1
                            anchors.rightMargin: AppTheme.spaceS
                            color: isSelected ? AppTheme.stateSelected : (menuHover.pressed ? AppTheme.statePress : (menuHover.containsMouse ? AppTheme.stateHover : "transparent"))
                            radius: AppTheme.radiusMedium
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        }

                        Row {
                            anchors.verticalCenter: parent.verticalCenter
                            anchors.left: parent.left
                            anchors.leftMargin: AppTheme.spaceL
                            spacing: AppTheme.spaceM

                            Rectangle {
                                width: 32
                                height: 32
                                radius: AppTheme.radiusSmall
                                color: isSelected ? AppTheme.accentBrand : AppTheme.bgBase
                                anchors.verticalCenter: parent.verticalCenter
                                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                                IconImage {
                                    anchors.centerIn: parent
                                    source: modelData.icon
                                    width: AppTheme.iconMedium
                                    height: AppTheme.iconMedium
                                    color: isSelected ? AppTheme.textOnAccent : AppTheme.textSecondary
                                }
                            }

                            Column {
                                anchors.verticalCenter: parent.verticalCenter
                                spacing: 1

                                Text {
                                    text: modelData.title
                                    color: isSelected ? AppTheme.accentBrand : AppTheme.textPrimary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeBody
                                    font.weight: isSelected ? AppTheme.weightBold : AppTheme.weightMedium
                                }

                                Text {
                                    text: modelData.desc
                                    color: AppTheme.textTertiary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeMicro
                                }
                            }
                        }

                        MouseArea {
                            id: menuHover
                            anchors.fill: parent
                            hoverEnabled: true
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                settingsMenuList.currentIndex = index
                                settingsStack.currentIndex = index
                            }
                        }
                    }
                }
            }
        }

        // ==========================================
        // ПРАВАЯ ОБЛАСТЬ КОНТЕНТА
        // ==========================================
        Item {
            width: parent.width - settingsMenuPanel.width
            height: parent.height

            Rectangle {
                anchors.fill: parent
                color: AppTheme.bgBase
            }

            StackLayout {
                id: settingsStack
                anchors.fill: parent
                currentIndex: settingsMenuList.currentIndex

                // ── 1. Внешний вид ────────────────────────
                SettingsPage {
                    title: "Внешний вид"
                    description: "Тема оформления и режим ввода времени."

                    AppSwitch {
                        text: "Тёмная тема"
                        checked: backend.isDarkTheme
                        // onToggled срабатывает только от клика человека —
                        // обновление от программы не зациклит переключение
                        onToggled: backend.toggleTheme()
                    }

                    AppComboBox {
                        width: 340
                        label: "Режим ввода времени дежурств:"
                        cutoutColor: AppTheme.bgBase
                        model: [
                            { text: "Горизонтальный ползунок (Слайдер)", value: "slider" },
                            { text: "Крутящиеся барабаны (Как в iOS)",   value: "tumbler" }
                        ]
                        textRole: "text"
                        valueRole: "value"
                        Component.onCompleted: currentIndex = backend.timeInputMode === "tumbler" ? 1 : 0
                        onActivated: function(index) { backend.setTimeInputMode(currentValue) }
                    }
                }

                // ── 2. Горячие клавиши ────────────────────
                HotkeySettingsPanel {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                }

                // ── 3. Управление базами ──────────────────
                Item {
                    FileDialog {
                        id: importFileDialog
                        title: "Выберите базу для импорта"
                        nameFilters: ["SQLite файлы (*.sqlite *.db)", "Все файлы (*)"]
                        onAccepted: backend.importDatabaseCopy(importFileDialog.selectedFile)
                    }
                    FolderDialog {
                        id: changeStorageDialog
                        title: "Выберите новую папку хранения"
                        property string pendingFolder: ""
                        onAccepted: {
                            pendingFolder = changeStorageDialog.selectedFolder
                            mainWindow.askConfirm(
                                "Перенести все базы?",
                                "Файлы всех подразделений будут физически перенесены в выбранную папку.",
                                "Перенести",
                                function() { backend.changeDbDirectory(changeStorageDialog.pendingFolder) },
                                false
                            )
                        }
                    }
                    FolderDialog {
                        id: exportFolderDialog
                        title: "Выберите папку для сохранения копии"
                        property string pathToExport: ""
                        onAccepted: backend.exportDatabaseCopy(pathToExport, exportFolderDialog.selectedFolder)
                    }

                    ScrollView {
                        id: dbScroll
                        anchors.fill: parent
                        clip: true
                        ScrollBar.vertical.policy: ScrollBar.AsNeeded
                        ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

                        ColumnLayout {
                            x: AppTheme.spaceXL
                            width: Math.max(0, dbScroll.availableWidth - AppTheme.spaceXL * 2)
                            spacing: AppTheme.spaceL

                            Column {
                                spacing: AppTheme.spaceXS
                                Layout.fillWidth: true

                                Text {
                                    text: "Управление базами данных"
                                    color: AppTheme.textPrimary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeH2
                                    font.weight: AppTheme.weightBold
                                }
                                Text {
                                    text: "Переключайтесь между отделами, создавайте новые и управляйте существующими базами."
                                    color: AppTheme.textSecondary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeBody
                                    wrapMode: Text.WordWrap
                                    width: parent.width
                                }
                            }

                            // Единая панель вместо отдельных карточек:
                            // один фон, внутри разделы с едва заметными разделителями.
                            Rectangle {
                                Layout.fillWidth: true
                                // Высота панели = высота содержимого + поля, чтобы поля
                                // не накладывались друг на друга.
                                Layout.preferredHeight: dbPanelContent.implicitHeight + AppTheme.spaceXL * 2
                                color: AppTheme.bgSurface
                                radius: AppTheme.radiusLarge
                                border.color: AppTheme.borderDivider
                                border.width: 1

                                ColumnLayout {
                                    id: dbPanelContent
                                    anchors.left: parent.left
                                    anchors.right: parent.right
                                    anchors.top: parent.top
                                    anchors.margins: AppTheme.spaceXL
                                    spacing: AppTheme.spaceM

                                    Text {
                                        text: "Подключённые базы"
                                        color: AppTheme.textSecondary
                                        font.family: AppTheme.fontFamily
                                        font.pixelSize: AppTheme.sizeSmall
                                        font.weight: AppTheme.weightBold
                                        font.letterSpacing: 0.8
                                    }

                                    Item {
                                        Layout.fillWidth: true
                                        // Минимум 120px — чтобы сообщение о пустом списке не пряталось
                                        implicitHeight: Math.max(120, dbSettingsList.contentHeight)

                                        Column {
                                            visible: backend.dbList.length === 0
                                            anchors.centerIn: parent
                                            spacing: AppTheme.spaceS

                                            IconImage {
                                                source: "../icons/database.svg"
                                                width: 40
                                                height: 40
                                                color: AppTheme.textTertiary
                                                anchors.horizontalCenter: parent.horizontalCenter
                                                opacity: 0.5
                                            }
                                            Text {
                                                text: "Нет подключённых баз данных"
                                                color: AppTheme.textTertiary
                                                font.family: AppTheme.fontFamily
                                                font.pixelSize: AppTheme.sizeBody
                                                anchors.horizontalCenter: parent.horizontalCenter
                                            }
                                        }

                                        ListView {
                                            id: dbSettingsList
                                            anchors.fill: parent
                                            spacing: AppTheme.spaceXS
                                            clip: true
                                            model: backend.dbList

                                            delegate: Rectangle {
                                                width: ListView.view.width
                                                height: 60
                                                radius: AppTheme.radiusMedium
                                                color: dbHoverArea.containsMouse ? AppTheme.stateHover : AppTheme.bgBase
                                                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                                                property bool isActive: backend.activeDepartmentName === modelData.name

                                                Rectangle {
                                                    visible: isActive
                                                    anchors.left: parent.left
                                                    anchors.top: parent.top
                                                    anchors.bottom: parent.bottom
                                                    width: 4
                                                    color: AppTheme.accentBrand
                                                    radius: 2
                                                }

                                                Row {
                                                    anchors.left: parent.left
                                                    anchors.leftMargin: isActive ? AppTheme.spaceM + 4 : AppTheme.spaceM
                                                    anchors.verticalCenter: parent.verticalCenter
                                                    spacing: AppTheme.spaceS
                                                    z: 1

                                                    Rectangle {
                                                        width: 36
                                                        height: 36
                                                        radius: AppTheme.radiusSmall
                                                        color: isActive ? AppTheme.accentBrand : AppTheme.bgSurface
                                                        anchors.verticalCenter: parent.verticalCenter

                                                        IconImage {
                                                            anchors.centerIn: parent
                                                            source: "../icons/database.svg"
                                                            width: AppTheme.iconMedium
                                                            height: AppTheme.iconMedium
                                                            color: isActive ? AppTheme.textOnAccent : AppTheme.textSecondary
                                                        }
                                                    }

                                                    Column {
                                                        anchors.verticalCenter: parent.verticalCenter
                                                        spacing: 2

                                                        Row {
                                                            spacing: AppTheme.spaceXS

                                                            Text {
                                                                text: modelData.name
                                                                color: AppTheme.textPrimary
                                                                font.family: AppTheme.fontFamily
                                                                font.pixelSize: AppTheme.sizeBody
                                                                font.weight: AppTheme.weightBold
                                                            }

                                                            Rectangle {
                                                                visible: isActive
                                                                height: 18
                                                                width: activeLabel.implicitWidth + 12
                                                                radius: 9
                                                                color: AppTheme.accentBrand
                                                                anchors.verticalCenter: parent.verticalCenter

                                                                Text {
                                                                    id: activeLabel
                                                                    anchors.centerIn: parent
                                                                    text: "Активна"
                                                                    color: AppTheme.textOnAccent
                                                                    font.family: AppTheme.fontFamily
                                                                    font.pixelSize: AppTheme.sizeMicro
                                                                    font.weight: AppTheme.weightBold
                                                                }
                                                            }
                                                        }

                                                        Text {
                                                            text: modelData.path
                                                            color: AppTheme.textTertiary
                                                            font.family: AppTheme.fontFamily
                                                            font.pixelSize: AppTheme.sizeMicro
                                                            width: 280
                                                            elide: Text.ElideMiddle
                                                        }
                                                    }
                                                }

                                                Row {
                                                    anchors.right: parent.right
                                                    anchors.rightMargin: AppTheme.spaceM
                                                    anchors.verticalCenter: parent.verticalCenter
                                                    spacing: AppTheme.spaceXS
                                                    z: 2

                                                    AppButton {
                                                        text: "Папка"
                                                        width: 80
                                                        variant: "secondary"
                                                        onClicked: backend.openDbFolder(modelData.path)
                                                    }
                                                    AppButton {
                                                        text: "Экспорт"
                                                        width: 80
                                                        variant: "secondary"
                                                        onClicked: {
                                                            exportFolderDialog.pathToExport = modelData.path
                                                            exportFolderDialog.open()
                                                        }
                                                    }
                                                    AppButton {
                                                        text: "Открыть"
                                                        width: 80
                                                        variant: "primary"
                                                        onClicked: {
                                                            backend.openDatabase(modelData.path)
                                                            root.close()
                                                        }
                                                    }
                                                    AppButton {
                                                        text: "Убрать"
                                                        width: 80
                                                        variant: "danger"
                                                        onClicked: {
                                                            let dbPath = modelData.path
                                                            mainWindow.askConfirm(
                                                                "Убрать базу из списка?",
                                                                "Подразделение «" + modelData.name + "» исчезнет из списка выбора.\nСам файл на диске удалён НЕ будет — его можно подключить обратно.",
                                                                "Убрать",
                                                                function() { backend.removeDatabaseFromList(dbPath) }
                                                            )
                                                        }
                                                    }
                                                }

                                                MouseArea {
                                                    id: dbHoverArea
                                                    anchors.fill: parent
                                                    hoverEnabled: true
                                                    propagateComposedEvents: true
                                                    onPressed: mouse.accepted = false
                                                }
                                            }
                                        }
                                    }

                                    // Разделитель разделов
                                    Rectangle {
                                        Layout.fillWidth: true
                                        height: 1
                                        color: AppTheme.borderDivider
                                        opacity: 0.6
                                    }

                                    Text {
                                        text: "Создать или подключить"
                                        color: AppTheme.textSecondary
                                        font.family: AppTheme.fontFamily
                                        font.pixelSize: AppTheme.sizeSmall
                                        font.weight: AppTheme.weightBold
                                        font.letterSpacing: 0.8
                                    }

                                    RowLayout {
                                        Layout.fillWidth: true
                                        spacing: AppTheme.spaceS

                                        AppTextField {
                                            id: newDbNameInput
                                            Layout.fillWidth: true
                                            placeholderText: "Название нового отдела..."
                                            cutoutColor: AppTheme.bgSurface
                                            Keys.onReturnPressed: {
                                                if (text.trim() !== "") {
                                                    backend.createNewDatabase(text)
                                                    text = ""
                                                }
                                            }
                                        }

                                        AppButton {
                                            text: "Создать базу"
                                            iconSource: "../icons/database.svg"
                                            Layout.preferredWidth: 140
                                            variant: "primary"
                                            onClicked: {
                                                backend.createNewDatabase(newDbNameInput.text)
                                                newDbNameInput.text = ""
                                            }
                                        }
                                    }

                                    RowLayout {
                                        Layout.fillWidth: true
                                        spacing: AppTheme.spaceS

                                        AppButton {
                                            text: "Подключить файл..."
                                            iconSource: "../icons/folder.svg"
                                            Layout.preferredWidth: 170
                                            variant: "secondary"
                                            onClicked: root.requestFileAttach()
                                        }
                                        AppButton {
                                            text: "Импорт копии..."
                                            iconSource: "../icons/folder.svg"
                                            Layout.preferredWidth: 160
                                            variant: "secondary"
                                            onClicked: importFileDialog.open()
                                        }
                                        AppButton {
                                            text: "Изменить путь хранения..."
                                            iconSource: "../icons/settings.svg"
                                            Layout.fillWidth: true
                                            variant: "secondary"
                                            onClicked: changeStorageDialog.open()
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // ── 4. Данные подразделения ───────────────
                Item {
                    ScrollView {
                        id: deptScroll
                        anchors.fill: parent
                        clip: true
                        ScrollBar.vertical.policy: ScrollBar.AsNeeded
                        ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

                        ColumnLayout {
                            x: AppTheme.spaceXL
                            width: Math.max(0, deptScroll.availableWidth - AppTheme.spaceXL * 2)
                            spacing: AppTheme.spaceL

                            Column {
                                spacing: AppTheme.spaceXS
                                Layout.fillWidth: true

                                Text {
                                    text: "Данные подразделения"
                                    color: AppTheme.textPrimary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeH2
                                    font.weight: AppTheme.weightBold
                                }
                                Text {
                                    text: "Эти данные используются в шапке и подписях при экспорте документов."
                                    color: AppTheme.textSecondary
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeBody
                                    wrapMode: Text.WordWrap
                                    width: parent.width
                                }
                            }

                            // Единая панель с разделами и разделителями
                            Rectangle {
                                Layout.fillWidth: true
                                // Высота панели = высота содержимого + поля, чтобы поля
                                // не накладывались друг на друга.
                                Layout.preferredHeight: deptPanelContent.implicitHeight + AppTheme.spaceXL * 2
                                color: AppTheme.bgSurface
                                radius: AppTheme.radiusLarge
                                border.color: AppTheme.borderDivider
                                border.width: 1

                                ColumnLayout {
                                    id: deptPanelContent
                                    anchors.left: parent.left
                                    anchors.right: parent.right
                                    anchors.top: parent.top
                                    anchors.margins: AppTheme.spaceXL
                                    spacing: AppTheme.spaceL

                                    SettingsFormSection {
                                        title: "Организация"
                                        Layout.fillWidth: true

                                        AppTextField {
                                            id: deptNameInput
                                            width: parent.width * 0.65
                                            label: "Название отдела:"
                                            text: backend.departmentData.department_name || ""
                                            cutoutColor: AppTheme.bgSurface
                                        }
                                    }

                                    Rectangle {
                                        Layout.fillWidth: true
                                        height: 1
                                        color: AppTheme.borderDivider
                                        opacity: 0.6
                                    }

                                    SettingsFormSection {
                                        title: "Ответственное лицо"
                                        Layout.fillWidth: true

                                        RowLayout {
                                            width: parent.width
                                            spacing: AppTheme.spaceM

                                            AppTextField {
                                                id: deptPosInput
                                                Layout.fillWidth: true
                                                label: "Должность:"
                                                text: backend.departmentData.resp_position || ""
                                                cutoutColor: AppTheme.bgSurface
                                            }
                                            AppTextField {
                                                id: deptRankInput
                                                Layout.preferredWidth: 180
                                                label: "Звание:"
                                                text: backend.departmentData.resp_rank || ""
                                                cutoutColor: AppTheme.bgSurface
                                            }
                                        }

                                        RowLayout {
                                            width: parent.width
                                            spacing: AppTheme.spaceM

                                            AppTextField {
                                                id: deptLastInput
                                                Layout.fillWidth: true
                                                label: "Фамилия:"
                                                text: backend.departmentData.resp_last_name || ""
                                                cutoutColor: AppTheme.bgSurface
                                            }
                                            AppTextField {
                                                id: deptFirstInput
                                                Layout.fillWidth: true
                                                label: "Имя:"
                                                text: backend.departmentData.resp_first_name || ""
                                                cutoutColor: AppTheme.bgSurface
                                            }
                                            AppTextField {
                                                id: deptMidInput
                                                Layout.fillWidth: true
                                                label: "Отчество:"
                                                text: backend.departmentData.resp_middle_name || ""
                                                cutoutColor: AppTheme.bgSurface
                                            }
                                        }
                                    }
                                }
                            }

                            RowLayout {
                                Layout.fillWidth: true

                                Item { Layout.fillWidth: true }

                                AppButton {
                                    text: "Сохранить изменения"
                                    iconSource: "../icons/check.svg"
                                    Layout.preferredWidth: 220
                                    variant: "primary"
                                    onClicked: backend.saveDepartmentData(
                                        deptNameInput.text,
                                        deptPosInput.text,
                                        deptRankInput.text,
                                        deptLastInput.text,
                                        deptFirstInput.text,
                                        deptMidInput.text
                                    )
                                }
                            }
                        }
                    }
                }

                // ── 5. Уведомления ────────────────────────
                SettingsPage {
                    title: "Уведомления"
                    description: "Настройте, как программа напоминает о важных делах."

                    Text {
                        width: parent.width
                        text: "Напоминание о сдаче табеля"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBodyLarge
                        font.weight: AppTheme.weightBold
                    }
                    Text {
                        width: parent.width
                        text: "С 28-го числа текущего месяца по 5-е число следующего программа напоминает подготовить табель, ознакомить с ним сотрудников и сдать его в кадровое подразделение. Напоминание появляется не чаще одного раза в 3 часа."
                        color: AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        wrapMode: Text.WordWrap
                    }

                    AppSwitch {
                        text: "Напоминать о сдаче табеля"
                        checked: backend.reminderEnabled
                        onCheckedChanged: backend.setReminderEnabled(checked)
                    }
                }

                // ── 6. Обновление ─────────────────────────
                Item {
                    FileDialog {
                        id: updateZipDialog
                        title: "Выберите архив новой версии"
                        nameFilters: ["Архивы OVERTIMETAB (*.zip)", "Все файлы (*)"]
                        onAccepted: backend.prepareUpdateFromPath(updateZipDialog.selectedFile)
                    }
                    FolderDialog {
                        id: updateFolderDialog
                        title: "Выберите папку с новой версией"
                        onAccepted: backend.prepareUpdateFromPath(updateFolderDialog.selectedFolder)
                    }

                    SettingsPage {
                        anchors.fill: parent
                        title: "Обновление"
                        description: "Базы, горячие клавиши и тема остаются на месте. Меняется только сама программа."

                        Text {
                            width: parent.width
                            text: "Сейчас стоит " + AppTheme.appVersionFull
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBodyLarge
                            font.weight: AppTheme.weightBold
                        }
                        Text {
                            width: parent.width
                            text: "Положите zip рядом с OVERTIMETAB.exe или на флешку — программа сама её заметит (имя файла не важно, смотрим содержимое) и покажет кнопку внизу слева. Либо укажите файл вручную."
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            wrapMode: Text.WordWrap
                        }

                        RowLayout {
                            width: parent.width
                            spacing: AppTheme.spaceS

                            AppButton {
                                text: "Указать архив .zip"
                                variant: "secondary"
                                Layout.preferredWidth: 200
                                enabled: !backend.updateBusy
                                onClicked: updateZipDialog.open()
                            }
                            AppButton {
                                text: "Указать папку"
                                variant: "secondary"
                                Layout.preferredWidth: 170
                                enabled: !backend.updateBusy
                                onClicked: updateFolderDialog.open()
                            }
                        }

                        AppButton {
                            visible: backend.updateReady
                            text: backend.updateVersion
                                  ? ("Перезапустить и обновить до " + backend.updateVersion)
                                  : "Перезапустить и обновить"
                            variant: "primary"
                            width: Math.min(parent.width, 420)
                            onClicked: backend.applyReadyUpdate()
                        }

                        Text {
                            visible: backend.updateBusy
                            width: parent.width
                            text: backend.updateStatusText || "Готовим обновление…"
                            color: AppTheme.accentBrand
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                        }
                    }
                }
            }
        }
    }
}
