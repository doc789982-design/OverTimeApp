import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl
import QtQuick.Window
import QtQuick.Dialogs
import QtQuick.Layouts
import Qt5Compat.GraphicalEffects
import "components" as AppUI 


ApplicationWindow {
    id: mainWindow 
    visible: !backend.startHidden
    width: 1200  
    height: 800
    // Посуда не бьётся: меньше этого размера окно не сожмёшь,
    // интерфейс не превратится в кашу
    minimumWidth: 940
    minimumHeight: 620
    title: "OVERTIMETAB " + AppUI.AppTheme.appVersionFull

    // Прозрачный фон самого окна
    color: AppUI.AppTheme.bgBase
    flags: Qt.Window | Qt.FramelessWindowHint

    property Item activeSpotlightCell: null

    // Колбэк для диалога подтверждения (хранится, пока пользователь думает)
    property var confirmCallback: null

    // ЕДИНАЯ ТОЧКА ВХОДА ДЛЯ ВСЕХ ПОДТВЕРЖДЕНИЙ В ПРОГРАММЕ.
    // Пример: mainWindow.askConfirm("Удалить?", "Точно?", "Удалить", function() { ... })
    function askConfirm(titleText, messageText, confirmText, callback, isDanger) {
        confirmDialog.dialogTitle = titleText
        confirmDialog.dialogMessage = messageText
        confirmDialog.confirmLabel = (confirmText && confirmText !== "") ? confirmText : "Подтвердить"
        confirmDialog.dangerMode = (isDanger === undefined) ? true : isDanger
        mainWindow.confirmCallback = callback
        confirmDialog.open()
    }

    // Сворачивание в трей: прячем окно и ОДИН раз объясняем, куда оно делось.
    // (Как в Telegram: «Приложение продолжит работу в фоне»)
    function minimizeToTray() {
        mainWindow.hide()
        if (!backend.trayHintWasShown()) {
            systemAlert.showCustom(
                "OVERTIMETAB продолжит работать в фоне. Развернуть — клик по иконке возле часов.",
                "Развернуть"
            )
            backend.setTrayHintShown()
        }
    }

    // Alt+F4 и системное закрытие окна тоже сворачивают в трей, а не убивают программу.
    // Полный выход — через иконку в трее: «Закрыть полностью».
    onClosing: (close) => {
        close.accepted = false
        minimizeToTray()
    }

    Connections {
        target: backend
        function onDatabaseOpened() { 
            stackView.replace(mainWorkspacePage) 
        }
        function onItemDeleted(dateStr, colorHex) {
            let workspace = stackView.currentItem
            if (workspace && workspace.calendarPanel) {
                workspace.calendarPanel.explodeDay(dateStr, colorHex)
            }
        }
    }

    // СЛОЙ 2: Основной визуальный корень
    // Item с clip: false — просто контейнер для позиционирования
    // СЛОЙ: Основной визуальный корень
    Item {
        id: visualRoot
        anchors.fill: parent

        Rectangle {
            id: mainBackground
            anchors.fill: parent

            color: AppUI.AppTheme.bgBase

            // Бордер поверх всего содержимого
            Rectangle {
                anchors.fill: parent
                color: "transparent"
                border.color: AppUI.AppTheme.borderDivider
                border.width: 1
                z: 999
            }

            Rectangle {
                id: customTitleBar
                anchors.top: parent.top
                anchors.left: parent.left
                anchors.right: parent.right
                height: 36
                color: AppUI.AppTheme.bgBase
                z: 100

                MouseArea {
                    anchors.fill: parent
                    anchors.rightMargin: 150
                    onPressed: mainWindow.startSystemMove()
                    onDoubleClicked: {
                        if (mainWindow.visibility === Window.Maximized) {
                            mainWindow.showNormal()
                        } else {
                            mainWindow.showMaximized()
                        }
                    }
                }

                Text {
                    anchors.left: parent.left
                    anchors.leftMargin: AppUI.AppTheme.spaceM
                    anchors.verticalCenter: parent.verticalCenter
                    text: "OVERTIMETAB " + AppUI.AppTheme.appVersionFull
                    color: AppUI.AppTheme.textTertiary
                    font.family: AppUI.AppTheme.fontFamily
                    font.pixelSize: AppUI.AppTheme.sizeSmall
                    font.weight: AppUI.AppTheme.weightBold
                    font.letterSpacing: 2
                }

                Row {
                    anchors.right: parent.right
                    anchors.top: parent.top
                    anchors.bottom: parent.bottom

                    Rectangle {
                        width: 46
                        height: parent.height
                        color: helpHov.pressed ? AppUI.AppTheme.statePress : (helpHov.containsMouse ? AppUI.AppTheme.stateHover : "transparent")
                        
                        MouseArea {
                            id: helpHov
                            anchors.fill: parent
                            hoverEnabled: true
                            cursorShape: Qt.PointingHandCursor
                            onClicked: helpDialog.show()
                        }
                        
                        IconImage {
                            anchors.centerIn: parent
                            source: "icons/help.svg"
                            width: AppUI.AppTheme.iconMedium
                            height: AppUI.AppTheme.iconMedium
                            color: helpHov.containsMouse ? AppUI.AppTheme.textPrimary : AppUI.AppTheme.textSecondary
                        }
                        
                        AppUI.AppToolTip {
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.top: parent.bottom 
                            anchors.topMargin: AppUI.AppTheme.spaceXXS       
                            dropDown: true             
                            isVisible: helpHov.containsMouse
                            text: "Справка и горячие клавиши (F1)"

                            // Список горячих клавиш из настроек
                            Rectangle {
                                visible: backend.hotkeysList.length > 0
                                width: 260
                                height: 1
                                color: AppUI.AppTheme.borderDivider
                            }

                            Column {
                                width: 260
                                spacing: AppUI.AppTheme.spaceXS
                                visible: backend.hotkeysList.length > 0

                                Text {
                                    text: "Горячие клавиши:"
                                    color: AppUI.AppTheme.textPrimary
                                    font.family: AppUI.AppTheme.fontFamily
                                    font.pixelSize: AppUI.AppTheme.sizeSmall
                                    font.weight: AppUI.AppTheme.weightBold
                                }

                                Repeater {
                                    model: backend.hotkeysList
                                    RowLayout {
                                        width: parent.width
                                        spacing: AppUI.AppTheme.spaceS

                                        Rectangle {
                                            Layout.preferredWidth: Math.max(44, hkKey.implicitWidth + 16)
                                            Layout.preferredHeight: 26
                                            radius: AppUI.AppTheme.radiusSmall
                                            color: AppUI.AppTheme.bgBase
                                            border.color: AppUI.AppTheme.borderInput
                                            border.width: 1
                                            Text {
                                                id: hkKey
                                                anchors.centerIn: parent
                                                text: modelData.key
                                                color: AppUI.AppTheme.accentBrand
                                                font.family: AppUI.AppTheme.fontFamily
                                                font.pixelSize: AppUI.AppTheme.sizeSmall
                                                font.weight: AppUI.AppTheme.weightBold
                                            }
                                        }

                                        Text {
                                            Layout.fillWidth: true
                                            text: helpDialog.describeHotkey(modelData)
                                            color: AppUI.AppTheme.textSecondary
                                            font.family: AppUI.AppTheme.fontFamily
                                            font.pixelSize: AppUI.AppTheme.sizeSmall
                                            wrapMode: Text.WordWrap
                                            elide: Text.ElideRight
                                            maximumLineCount: 2
                                        }
                                    }
                                }
                            }
                        }
                    }

                    Rectangle {
                        width: 46
                        height: parent.height
                        color: minHov.pressed ? AppUI.AppTheme.statePress : (minHov.containsMouse ? AppUI.AppTheme.stateHover : "transparent")
                        Rectangle {
                            anchors.centerIn: parent
                            width: 10; height: 1
                            color: AppUI.AppTheme.textSecondary
                        }
                        MouseArea {
                            id: minHov
                            anchors.fill: parent
                            hoverEnabled: true
                            onClicked: mainWindow.showMinimized()
                        }
                    }

                    Rectangle {
                        width: 46
                        height: parent.height
                        color: maxHov.pressed ? AppUI.AppTheme.statePress : (maxHov.containsMouse ? AppUI.AppTheme.stateHover : "transparent")
                        Rectangle {
                            anchors.centerIn: parent
                            width: 10; height: 10
                            color: "transparent"
                            border.color: AppUI.AppTheme.textSecondary
                            border.width: 1
                        }
                        MouseArea {
                            id: maxHov
                            anchors.fill: parent
                            hoverEnabled: true
                            onClicked: {
                                if (mainWindow.visibility === Window.Maximized) {
                                    mainWindow.showNormal()
                                } else {
                                    mainWindow.showMaximized()
                                }
                            }
                        }
                    }

                    Rectangle {
                        width: 46
                        height: parent.height
                        color: closeHov.pressed ? Qt.darker(AppUI.AppTheme.accentDanger, 1.2) : (closeHov.containsMouse ? AppUI.AppTheme.accentDanger : "transparent")
                        Text {
                            anchors.centerIn: parent
                            text: "✕"
                            color: closeHov.containsMouse ? AppUI.AppTheme.textOnAccent : AppUI.AppTheme.textSecondary
                            font.pixelSize: AppUI.AppTheme.sizeSmall
                            font.weight: AppUI.AppTheme.weightBold
                        }
                        MouseArea {
                            id: closeHov
                            anchors.fill: parent
                            hoverEnabled: true
                            onClicked: minimizeToTray()
                        }
                    }
                }
            }

            MouseArea {
                z: 200; width: 8
                anchors.left: parent.left
                anchors.top: parent.top
                anchors.bottom: parent.bottom
                cursorShape: Qt.SizeHorCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.LeftEdge)
            }
            MouseArea {
                z: 200; width: 8
                anchors.right: parent.right
                anchors.top: parent.top
                anchors.bottom: parent.bottom
                cursorShape: Qt.SizeHorCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.RightEdge)
            }
            MouseArea {
                z: 200; height: 8
                anchors.top: parent.top
                anchors.left: parent.left
                anchors.right: parent.right
                cursorShape: Qt.SizeVerCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.TopEdge)
            }
            MouseArea {
                z: 200; height: 8
                anchors.bottom: parent.bottom
                anchors.left: parent.left
                anchors.right: parent.right
                cursorShape: Qt.SizeVerCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.BottomEdge)
            }

            // УГЛЫ: тянуть окно можно и за углы, как в обычных программах
            MouseArea {
                z: 201; width: 16; height: 16
                anchors.top: parent.top; anchors.left: parent.left
                cursorShape: Qt.SizeFDiagCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.TopEdge | Qt.LeftEdge)
            }
            MouseArea {
                z: 201; width: 16; height: 16
                anchors.top: parent.top; anchors.right: parent.right
                cursorShape: Qt.SizeBDiagCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.TopEdge | Qt.RightEdge)
            }
            MouseArea {
                z: 201; width: 16; height: 16
                anchors.bottom: parent.bottom; anchors.left: parent.left
                cursorShape: Qt.SizeBDiagCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.BottomEdge | Qt.LeftEdge)
            }
            MouseArea {
                z: 201; width: 16; height: 16
                anchors.bottom: parent.bottom; anchors.right: parent.right
                cursorShape: Qt.SizeFDiagCursor
                visible: mainWindow.visibility !== Window.Maximized
                onPressed: mainWindow.startSystemResize(Qt.BottomEdge | Qt.RightEdge)
            }

            StackView {
                id: stackView
                anchors.top: customTitleBar.bottom
                anchors.bottom: parent.bottom
                anchors.left: parent.left
                anchors.right: parent.right
                initialItem: dbSelectionPage

                replaceEnter: Transition {
                    NumberAnimation {
                        property: "opacity"
                        from: 0.0; to: 1.0
                        duration: 400
                        easing.type: Easing.OutSine
                    }
                }
                replaceExit: Transition {
                    NumberAnimation {
                        property: "opacity"
                        from: 1.0; to: 0.0
                        duration: 400
                        easing.type: Easing.OutSine
                    }
                }
            }
        }
    }

    // ==========================================
    // СТРАНИЦА ВЫБОРА БД
    // ==========================================
    Component {
        id: dbSelectionPage
        Item {
            id: dbPageRoot
            property string targetDbPath: ""
            Component.onCompleted: { 
                if (backend.dbList.length === 1) { 
                    targetDbPath = backend.dbList[0].path
                    autoStartTimer.start() 
                } 
            }
            Timer { 
                id: autoStartTimer
                interval: 100
                onTriggered: { 
                    rightContentArea.opacity = 0
                    expandAnim.start() 
                } 
            }
            function openDbAnimated(path) { 
                if (expandAnim.running) return
                targetDbPath = path
                rightContentArea.opacity = 0
                expandAnim.start() 
            }
            
            Item {
                id: rightContentArea
                anchors.fill: parent
                anchors.leftMargin: 300 
                Behavior on opacity { 
                    NumberAnimation { duration: 250; easing.type: Easing.OutQuad } 
                }
                Text { 
                    id: titleText
                    text: "Выберите подразделение"
                    color: AppUI.AppTheme.textPrimary
                    font.pixelSize: AppUI.AppTheme.sizeH1
                    font.weight: AppUI.AppTheme.weightBold
                    anchors.top: parent.top
                    anchors.topMargin: AppUI.AppTheme.spaceXXL
                    anchors.left: parent.left
                    anchors.leftMargin: AppUI.AppTheme.spaceXXL 
                }
                ListView {
                    id: dbListView
                    anchors.top: titleText.bottom
                    anchors.topMargin: AppUI.AppTheme.spaceXL
                    anchors.bottom: bottomBar.top
                    anchors.bottomMargin: AppUI.AppTheme.spaceL
                    anchors.left: parent.left
                    anchors.right: parent.right
                    anchors.leftMargin: AppUI.AppTheme.spaceXXL
                    anchors.rightMargin: AppUI.AppTheme.spaceXXL
                    spacing: AppUI.AppTheme.spaceM
                    clip: true
                    model: backend.dbList
                    delegate: Rectangle {
                        width: ListView.view.width
                        height: AppUI.AppTheme.startCardHeight
                        radius: AppUI.AppTheme.radiusLarge
                        color: AppUI.AppTheme.bgElevated
                        border.color: AppUI.AppTheme.borderDivider
                        border.width: 1
                        Rectangle {
                            anchors.fill: parent
                            radius: parent.radius
                            color: mouseArea.pressed
                                   ? AppUI.AppTheme.statePress
                                   : (mouseArea.containsMouse ? AppUI.AppTheme.stateHover : "transparent")
                        }
                        Rectangle {
                            id: iconRect
                            width: 44; height: 44
                            radius: AppUI.AppTheme.radiusPill
                            color: AppUI.AppTheme.bgBrandSoft
                            anchors.left: parent.left
                            anchors.leftMargin: AppUI.AppTheme.spaceM
                            anchors.verticalCenter: parent.verticalCenter
                            Text {
                                anchors.centerIn: parent
                                text: modelData.name.charAt(0).toUpperCase()
                                color: AppUI.AppTheme.accentBrand
                                font.weight: AppUI.AppTheme.weightBold
                                font.pixelSize: AppUI.AppTheme.sizeH4
                            }
                        }
                        Column {
                            anchors.left: iconRect.right
                            anchors.leftMargin: AppUI.AppTheme.spaceM
                            anchors.right: folderBtn.left
                            anchors.rightMargin: AppUI.AppTheme.spaceS
                            anchors.verticalCenter: parent.verticalCenter
                            spacing: 2
                            Text {
                                width: parent.width
                                text: modelData.name
                                color: AppUI.AppTheme.textPrimary
                                font.pixelSize: AppUI.AppTheme.sizeBodyLarge
                                font.weight: AppUI.AppTheme.weightBold
                                elide: Text.ElideRight
                            }
                            Text {
                                width: parent.width
                                text: modelData.path
                                color: AppUI.AppTheme.textSecondary
                                font.pixelSize: AppUI.AppTheme.sizeSmall
                                elide: Text.ElideMiddle
                            }
                        }
                        Rectangle {
                            id: folderBtn
                            width: 36; height: 36
                            radius: AppUI.AppTheme.radiusPill
                            anchors.right: parent.right
                            anchors.rightMargin: AppUI.AppTheme.spaceM
                            anchors.verticalCenter: parent.verticalCenter
                            color: folderHov.pressed
                                   ? AppUI.AppTheme.statePress
                                   : (folderHov.containsMouse ? AppUI.AppTheme.stateHover : "transparent")
                            IconImage {
                                anchors.centerIn: parent
                                source: "icons/folder.svg"
                                width: AppUI.AppTheme.iconMedium
                                height: AppUI.AppTheme.iconMedium
                                color: AppUI.AppTheme.textSecondary
                            }
                            MouseArea {
                                id: folderHov
                                anchors.fill: parent
                                hoverEnabled: true
                                onClicked: backend.openDbFolder(modelData.path)
                            }
                        }
                        MouseArea {
                            id: mouseArea
                            anchors.fill: parent
                            hoverEnabled: true
                            cursorShape: Qt.PointingHandCursor
                            onClicked: dbPageRoot.openDbAnimated(modelData.path)
                        }
                    }
                }
                Item {
                    id: bottomBar
                    height: 80
                    anchors.bottom: parent.bottom
                    anchors.left: parent.left
                    anchors.right: parent.right
                    anchors.leftMargin: AppUI.AppTheme.spaceXXL
                    anchors.rightMargin: AppUI.AppTheme.spaceXXL
                    Row {
                        anchors.verticalCenter: parent.verticalCenter
                        spacing: AppUI.AppTheme.spaceM
                        AppUI.AppButton {
                            width: 200
                            variant: "primary"
                            text: "Создать подразделение"
                            onClicked: createDbDialog.showCentered()
                        }
                        AppUI.AppButton {
                            width: 330
                            variant: "secondary"
                            text: "Подключить базу данных"
                            onClicked: fileDialog.open()
                        }
                        AppUI.AppButton {
                            width: 330
                            variant: "secondary"
                            text: "Перенести базы в другую папку"
                            onClicked: globalFolderDialog.open()
                        }
                    }
                }
                FolderDialog {
                    id: globalFolderDialog
                    title: "Выберите новую папку для баз"
                    property string pendingFolder: ""
                    onAccepted: {
                        pendingFolder = globalFolderDialog.selectedFolder
                        mainWindow.askConfirm(
                            "Перенести все базы?",
                            "Файлы всех подразделений будут физически перенесены в выбранную папку. Пара секунд — и готово.",
                            "Перенести",
                            function() { backend.changeDbDirectory(globalFolderDialog.pendingFolder) },
                            false
                        )
                    }
                }
                FileDialog {
                    id: fileDialog
                    title: "Выберите базу"
                    nameFilters: ["SQLite файлы (*.sqlite *.db)", "Все файлы (*)"]
                    onAccepted: backend.attachDatabase(fileDialog.selectedFile)
                }
            }

            Rectangle { 
                id: leftPanel
                width: 300
                anchors.left: parent.left
                anchors.top: parent.top
                anchors.bottom: parent.bottom
                color: AppUI.AppTheme.bgPanel
                z: 10
                Rectangle {
                    anchors.right: parent.right
                    width: 1; height: parent.height
                    color: AppUI.AppTheme.borderDivider
                }
                Column {
                    anchors.centerIn: parent
                    width: 260
                    spacing: AppUI.AppTheme.spaceL
                    AppUI.AppEmptyMascot {
                        anchors.horizontalCenter: parent.horizontalCenter
                        width: 180; height: 180
                    }
                    Column {
                        id: logoText
                        width: parent.width
                        spacing: AppUI.AppTheme.spaceXXS
                        Text {
                            width: parent.width
                            text: "OVERTIMETAB"
                            color: AppUI.AppTheme.accentBrand
                            font.family: AppUI.AppTheme.fontFamily
                            font.pixelSize: AppUI.AppTheme.sizeH4
                            font.weight: AppUI.AppTheme.weightBold
                            font.letterSpacing: 1
                            horizontalAlignment: Text.AlignHCenter
                        }
                        Text {
                            width: parent.width
                            text: AppUI.AppTheme.appVersionFull
                            color: AppUI.AppTheme.textTertiary
                            font.family: AppUI.AppTheme.fontFamily
                            font.pixelSize: AppUI.AppTheme.sizeSmall
                            font.weight: AppUI.AppTheme.weightMedium
                            horizontalAlignment: Text.AlignHCenter
                        }
                    }
                }
                Column {
                    anchors.left: parent.left
                    anchors.right: parent.right
                    anchors.bottom: parent.bottom
                    anchors.bottomMargin: AppUI.AppTheme.spaceS
                    AppUI.UpdateBanner { width: parent.width }
                }
            }

            SequentialAnimation {
                id: expandAnim
                PauseAnimation { duration: 150 }
                NumberAnimation {
                    target: leftPanel; property: "width"
                    to: dbPageRoot.width; duration: 600
                    easing.type: Easing.InOutExpo
                }
                ScriptAction { script: backend.openDatabase(dbPageRoot.targetDbPath) }
            }
        }
    }

    // ==========================================
    // ГЛАВНАЯ СТРАНИЦА РАБОЧЕГО ПРОСТРАНСТВА
    // ==========================================
    Component {
        id: mainWorkspacePage
        Item {
            id: workspaceRoot
            property alias calendarPanel: calendarPanelId
            MouseArea {
                anchors.fill: parent
                z: 1
                propagateComposedEvents: true
                onPressed: (mouse) => {
                    empListPanel.dismissSearchIfOutside(this, mouse.x, mouse.y)
                    mouse.accepted = false
                }
            }
            SplitView {
                anchors.fill: parent
                z: 0
                orientation: Qt.Horizontal
                handle: Rectangle {
                    // Невидим по умолчанию; при наведении проявляется полусерая
                    // линия-«ручка», за которую можно потянуть и менять ширину.
                    implicitWidth: 8
                    color: "transparent"
                    opacity: (SplitHandle.hovered || SplitHandle.pressed) ? 1.0 : 0.0
                    Behavior on opacity { NumberAnimation { duration: 150 } }
                    Rectangle {
                        anchors.centerIn: parent
                        width: 3
                        height: parent.height
                        color: AppUI.AppTheme.borderInput
                    }
                }
                Item {
                    SplitView.preferredWidth: 350
                    SplitView.minimumWidth: 250
                    SplitView.maximumWidth: 600
                    SplitView {
                        anchors.fill: parent
                        orientation: Qt.Horizontal
                        handle: Rectangle {
                            implicitWidth: 1
                            color: AppUI.AppTheme.borderDivider
                            opacity: 0.4
                        }
                        AppUI.GroupSidebar {
                            SplitView.preferredWidth: 72
                            SplitView.minimumWidth: 72
                            SplitView.maximumWidth: 72
                            workspace: workspaceRoot
                        }
                        AppUI.EmployeeListPanel {
                            id: empListPanel
                            SplitView.fillWidth: true
                            workspace: workspaceRoot
                        }
                    }
                    Column {
                        anchors.left: parent.left
                        anchors.right: parent.right
                        anchors.bottom: parent.bottom
                        z: AppUI.AppTheme.zSticky
                        AppUI.UpdateBanner { width: parent.width }
                        AppUI.LeftControlPanel { width: parent.width }
                    }
                }
                AppUI.CalendarWorkspace {
                    id: calendarPanelId
                    SplitView.fillWidth: true
                }
            }
        }
    }

    // ==========================================
    // ОВЕРЕИ И ЭФФЕКТЫ
    // ==========================================
    // Размытие фона за окном убрали: окно просто открывается.
    // (Оверлей оставлен, но выключен, чтобы не трогать остальную механику.)
    AppUI.SpotlightOverlay {
        id: spotlightOverlay
        backgroundSource: visualRoot
        targetItem: mainWindow.activeSpotlightCell
        isActive: false
    }
    AppUI.ThemeTransition { id: themeTransition; targetItem: visualRoot }

    // Меню дня: открывается ЛКМ и ПКМ по ячейке календаря (как контекстное меню)
    AppUI.AppDayMenu { id: dayMenu }

    property var pendingBackendCall: null

    Item {
        id: thanosContainer
        anchors.fill: parent
        z: AppUI.AppTheme.zEffect
        Repeater { 
            id: thanosPool
            model: 10
            AppUI.ThanosEffect { 
                onSnapshotTaken: { 
                    if (mainWindow.pendingBackendCall) { 
                        mainWindow.pendingBackendCall()
                        mainWindow.pendingBackendCall = null 
                    } 
                } 
            } 
        }
    }

    function explodeMulti(itemsArray, backendCall) {
        if (!itemsArray || itemsArray.length === 0) { 
            backendCall()
            return 
        }
        mainWindow.pendingBackendCall = backendCall
        let explodedAny = false
        for (let i = 0; i < itemsArray.length; i++) {
            let item = itemsArray[i]
            if (item && item.visible && item.opacity > 0) {
                let t = null
                for (let j = 0; j < thanosPool.count; j++) { 
                    if (!thanosPool.itemAt(j).isExploding) { 
                        t = thanosPool.itemAt(j)
                        break 
                    } 
                }
                if (t) { 
                    t.explode(item)
                    explodedAny = true 
                }
            }
        }
        if (!explodedAny) { 
            mainWindow.pendingBackendCall = null
            backendCall() 
        }
    }

    function explodeAndDelete(dateStr, targetType, specificId, backendCall) {
        let workspace = stackView.currentItem
        if (!workspace || !workspace.calendarPanel) { 
            backendCall()
            return 
        }
        let itemsToExplode = []
        let cell = workspace.calendarPanel.findDayCell(dateStr)
        if (targetType === "duty") {
            let idsToExplode = []
            if (specificId !== null && specificId !== undefined) { 
                idsToExplode.push(specificId) 
            } else { 
                idsToExplode = workspace.calendarPanel.getDutyIdsInDay(dateStr) 
            }
            itemsToExplode = workspace.calendarPanel.findDutyRectsByIds(idsToExplode)
        } else if (targetType === "all") {
            let idsToExplode = workspace.calendarPanel.getDutyIdsInDay(dateStr)
            itemsToExplode = workspace.calendarPanel.findDutyRectsByIds(idsToExplode)
            if (cell) {
                if (cell.statusBarItemRef && cell.statusBarItemRef.visible)
                    itemsToExplode.push(cell.statusBarItemRef)
                if (cell.compBadgeItemRef && cell.compBadgeItemRef.visible)
                    itemsToExplode.push(cell.compBadgeItemRef)
            }
        } else {
            if (cell) {
                if (targetType === "status") itemsToExplode.push(cell.statusBarItemRef)
                else if (targetType === "comp") itemsToExplode.push(cell.compBadgeItemRef)
            }
        }
        explodeMulti(itemsToExplode, backendCall)
    }

    // ==========================================
    // ДИАЛОГИ
    // ==========================================
    AppUI.WhatsNewDialog    { id: whatsNewDialog }
    AppUI.HelpDialog        { id: helpDialog }
    AppUI.DayInspector      { id: dayInspector }
    AppUI.MoneyInspector    { 
        id: moneyInspector
        onRequestAddMoneyDialog: moneyDialog.openNew(moneyInspector.contentItem, 20, 20) 
    }
    AppUI.DayDutyDialog     { id: dayDutyDialog }
    AppUI.DayCompDialog     { id: dayCompDialog }
    AppUI.EmpDialog         { id: empDialog }
    AppUI.EndStatusDialog   { id: endStatusDialog }
    AppUI.TransferDialog    { id: transferDialog }
    AppUI.MoneyDialog       { id: moneyDialog }

    AppUI.AppConfirmDialog  {
        id: confirmDialog
        onAccepted: {
            let cb = mainWindow.confirmCallback
            mainWindow.confirmCallback = null
            if (cb) cb()
        }
        onRejected: mainWindow.confirmCallback = null
    }
    AppUI.SettingsDialog    { id: settingsDialog; onRequestFileAttach: fileDialog.open() }
    AppUI.AddGroupDialog    { id: addGroupDialog }
    AppUI.HistoryDialog     { id: historyDialog }
    AppUI.CreateDbDialog    { id: createDbDialog }

    // Тосты в собственном всегда-поверх-всего окне: всплывают над любыми
    // модальными окнами (настройки, печать и т.п.), а не за ними.
    AppUI.ToastWindow { id: toastWindow }

FileDialog { 
    id: exportDialog
    title: "Сохранить табель в Excel"
    fileMode: FileDialog.SaveFile
    nameFilters: ["Excel файлы (*.xlsx)"]
    currentFile: {
        let months = ["январь", "февраль", "март", "апрель", "май", "июнь", 
                      "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
        
        // Парсим backend.currentPeriodText (формат: "Январь 2026")
        let parts = backend.currentPeriodText.split(" ")
        let monthName = parts[0].toLowerCase() // "январь"
        let year = parts[1] // "2026"
        
        let deptName = backend.activeDepartmentName || "подразделение"
        
        return "Табель " + deptName + " за " + monthName + " " + year + ".xlsx"
    }
    onAccepted: backend.exportToExcel(exportDialog.selectedFile) 
}

    Shortcut { sequence: "Ctrl+Z";       onActivated: backend.undoAction() }
    Shortcut { sequence: "Ctrl+Y";       onActivated: backend.redoAction() }
    Shortcut { sequence: "Ctrl+Shift+Z"; onActivated: backend.redoAction() }
    Shortcut { sequence: "F1";           onActivated: helpDialog.show() }

    AppUI.AppPrintDialog { 
        id: customPrintDialog
        onPrintRequested: function(printerName, copies, pageFrom, pageTo, orientation, paperSize, collate) { 
            backend.quickPrint(printerName, copies, pageFrom, pageTo, orientation, paperSize, collate) 
        } 
    }

    AppUI.AppDesktopNotification {
        id: systemAlert
        // Кнопка «Развернуть» в подсказке про трей возвращает окно
        onActionTriggered: { mainWindow.show(); mainWindow.requestActivate() }
    }

    Timer {
        id: notificationTimer
        interval: 10800000; running: true; repeat: true
        onTriggered: { if (backend.reminderEnabled) checkAndNotify() }
    }
    Timer {
        id: startupNotificationTimer
        interval: 5000; running: true; repeat: false
        onTriggered: { if (backend.reminderEnabled) checkAndNotify() }
    }
    Timer {
        id: updateScanTimer
        // Редкая фоновая проверка. Частый опрос (5 сек) был незаметен-мигал
        // полоской и дёргал сеть; теперь — раз в 30 минут, и только при старте.
        interval: 1800000; running: true; repeat: true
        onTriggered: backend.scanForUpdates()
    }
    Timer {
        id: whatsNewTimer
        interval: 700; running: true; repeat: false
        onTriggered: whatsNewDialog.showIfNeeded()
    }

    function checkAndNotify() {
        let d = new Date()
        let day = d.getDate()
        if (day >= 28 || day <= 5) {
            systemAlert.showNotification()
        }
    }
}
