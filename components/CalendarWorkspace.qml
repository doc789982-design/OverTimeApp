import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects
import "."

Item {
    id: root
    
    property bool isYearView: false 
    
    // ═══════════════════════════════════════════════════════════════
    // ОТСЛЕЖИВАНИЕ СМЕНЫ СОТРУДНИКА/МЕСЯЦА
    // Анимация запускается ТОЛЬКО при этих изменениях
    // ═══════════════════════════════════════════════════════════════
    property int lastEmployeeId: -1
    property string lastPeriod: ""
    property bool needAnimation: true

    // ==========================================
    // ФУНКЦИИ ДЛЯ ЭФФЕКТА ТАНОСА
    // ==========================================
    function findDayCell(dateStr) {
        if (calendarStack.currentIndex !== 0) return null
        for (let i = 0; i < calendarGrid.children.length; i++) {
            let cell = calendarGrid.children[i]
            if (cell.dayInfo && cell.dayInfo.date_str === dateStr) return cell
        }
        return null
    }

    function getDutyIdsInDay(dateStr) {
        let cell = findDayCell(dateStr)
        let ids = []
        if (!cell || !cell.dutyColItemRef) return ids
        let children = cell.dutyColItemRef.children
        for (let i = 0; i < children.length; i++) {
            if (children[i].dutyId !== undefined) ids.push(children[i].dutyId)
        }
        return ids
    }

    function findDutyRectsByIds(idsArray) {
        let rects = []
        if (calendarStack.currentIndex !== 0) return rects
        for (let i = 0; i < calendarGrid.children.length; i++) {
            let cell = calendarGrid.children[i]
            if (cell.dutyColItemRef) {
                let children = cell.dutyColItemRef.children
                for (let j = 0; j < children.length; j++) {
                    if (children[j].dutyId !== undefined && idsArray.includes(children[j].dutyId)) {
                        rects.push(children[j])
                    }
                }
            }
        }
        return rects
    }

    // ═══════════════════════════════════════════════════════════════
    // ЗАПУСК АНИМАЦИИ (только при смене сотрудника/месяца)
    // ═══════════════════════════════════════════════════════════════
    function playMonthEnterAnimation() {
        for (let i = 0; i < calendarGrid.children.length; i++) {
            let cell = calendarGrid.children[i]
            if (cell && cell.cellAnim) {
                // Stagger: строки появляются с задержкой 40мс
                let rowDelay = Math.floor(i / 7) * 40
                cell.cellAnim.delay = rowDelay
                cell.cellAnim.resetAndEnter()
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════
    // ПРОВЕРКА: ИЗМЕНИЛСЯ ЛИ СОТРУДНИК ИЛИ ПЕРИОД
    // ═══════════════════════════════════════════════════════════════
    function checkAndAnimate() {
        let currentEmpId = backend.selectedEmployeeId
        let currentPeriod = backend.currentPeriodText
        
        let employeeChanged = (root.lastEmployeeId !== currentEmpId)
        let periodChanged = (root.lastPeriod !== currentPeriod)
        
        root.lastEmployeeId = currentEmpId
        root.lastPeriod = currentPeriod
        
        // Анимация только если сменился сотрудник или период
        if (employeeChanged || periodChanged) {
            root.needAnimation = true
            playMonthEnterAnimation()
        } else {
            root.needAnimation = false
        }
    }

    // ==========================================
    // ШАПКА КАЛЕНДАРЯ (Табы периодов)
    // ==========================================
    AppPeriodTabs {
        id: calendarHeader
        anchors.top: parent.top
        anchors.left: parent.left
        anchors.right: parent.right
        
        isYearView: root.isYearView
        
        onYearChanged: function(newYear) { backend.setYear(newYear) }
        
        onMonthClicked: function(month) { 
            root.isYearView = false
            calendarStack.currentIndex = 0
            backend.jumpToMonth(month) 
        }
        
        onYearViewClicked: { 
            root.isYearView = true
            calendarStack.currentIndex = 1 
        }
    }

    // ==========================================
    // ПАНЕЛЬ ИТОГОВ
    // ==========================================
    AppSummaryPanel {
        id: unifiedSummaryPanel
        isYearView: root.isYearView
        
        anchors.bottom: parent.bottom
        anchors.left: parent.left
        anchors.right: parent.right
        anchors.leftMargin: AppTheme.spaceL
        anchors.rightMargin: AppTheme.spaceL
        anchors.bottomMargin: AppTheme.spaceL
    }

    // ═══════════════════════════════════════════════════════════════
    // ПОДКЛЮЧЕНИЕ К СИГНАЛАМ
    // ═══════════════════════════════════════════════════════════════
    Connections {
        target: backend
        function onCalendarDaysChanged() {
            root.checkAndAnimate()
        }
    }

    StackLayout {
        id: calendarStack
        anchors.top: calendarHeader.bottom
        anchors.bottom: unifiedSummaryPanel.top 
        anchors.bottomMargin: AppTheme.spaceM
        anchors.left: parent.left
        anchors.right: parent.right
        currentIndex: 0

        // ==========================================
        // СТРАНИЦА 0: РЕЖИМ МЕСЯЦА
        // ==========================================
        Item {
            id: monthPage

            property string menuDate: ""
            property bool menuIsWeekend: false
            property bool menuIsHoliday: false
            property bool menuHasDuties: false
            property bool menuHasComps: false
            property var menuDayCell: null

            // ДНИ НЕДЕЛИ
            Row {
                id: weekDaysRow
                anchors.top: parent.top
                anchors.left: parent.left
                anchors.right: parent.right
                anchors.leftMargin: AppTheme.spaceL
                anchors.rightMargin: AppTheme.spaceL
                height: 40 
                
                Repeater {
                    model: ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
                    
                    Item { 
                        width: weekDaysRow.width / 7
                        height: parent.height
                        
                        Text { 
                            anchors.centerIn: parent
                            text: modelData
                            color: (index >= 5) ? AppTheme.accentDanger : AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightMedium 
                        } 
                    }
                }
            }

            // МАСКОТ
            Item {
                anchors.top: weekDaysRow.bottom
                anchors.bottom: parent.bottom
                anchors.left: parent.left
                anchors.right: parent.right
                visible: backend.selectedEmployeeId === 0
                z: 50

                Column {
                    anchors.centerIn: parent
                    spacing: AppTheme.spaceL

                    AppEmptyMascot {
                        anchors.horizontalCenter: parent.horizontalCenter
                    }

                    Column {
                        spacing: AppTheme.spaceXS
                        anchors.horizontalCenter: parent.horizontalCenter
                        
                        Text { 
                            text: "Сотрудник не выбран"
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeH1 
                            font.weight: AppTheme.weightBold
                            anchors.horizontalCenter: parent.horizontalCenter
                        } 
                        
                        Text { 
                            text: "Выберите карточку в панели слева, чтобы открыть табель"
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody 
                            anchors.horizontalCenter: parent.horizontalCenter
                        } 
                    }
                }
            }

            // СЕТКА КАЛЕНДАРЯ
            GridLayout {
                id: calendarGrid
                visible: backend.selectedEmployeeId !== 0 
                anchors.top: weekDaysRow.bottom
                anchors.bottom: parent.bottom 
                anchors.left: parent.left
                anchors.right: parent.right
                anchors.leftMargin: AppTheme.spaceL
                anchors.rightMargin: AppTheme.spaceL
                
                columns: 7
                rowSpacing: AppTheme.spaceXXS
                columnSpacing: AppTheme.spaceXXS

                Repeater {
                    model: 42 
                    
                    delegate: Item {
                        id: dayCell
                        Layout.fillWidth: true
                        Layout.fillHeight: true
                        
                        property alias mainContent: innerRect
                        property alias statusBarItemRef: statusBarItem
                        property alias compBadgeItemRef: compBadgeItem
                        property alias dutyColItemRef: dutyColItem
                        property alias cellAnim: dayCellAnim
                        
                        property var dayInfo: backend.calendarDays.length > index
                                              ? backend.calendarDays[index]
                                              : null
                        property bool isValid: dayInfo !== null && dayInfo.date_str !== undefined
                        
                        opacity: (isValid && dayInfo.is_current_month) ? 1.0 : 0.0
                        
                        property bool isToday: {
                            if (!isValid) return false
                            let d = new Date()
                            return dayInfo.date_str === (
                                d.getFullYear() + "-" +
                                ("0" + (d.getMonth() + 1)).slice(-2) + "-" +
                                ("0" + d.getDate()).slice(-2)
                            )
                        }

                        // Анимация появления ячейки
                        AppEnterAnimation {
                            id: dayCellAnim
                            target: innerRect
                        }

                        Rectangle {
                            id: innerRect
                            anchors.fill: parent
                            radius: AppTheme.radiusMedium
                            clip: true
                            
                            color: isValid && (dayInfo.is_weekend || dayInfo.is_holiday)
                                   ? AppTheme.bgDangerSoft
                                   : (isValid ? AppTheme.bgCell : "transparent")
                                   
                            scale: dayMouseArea.pressed ? AppTheme.scaleActive : 1.0
                            Behavior on scale {
                                NumberAnimation {
                                    duration: AppTheme.durMicro
                                    easing.type: AppTheme.easeStandard
                                }
                            }
                            Behavior on color {
                                ColorAnimation { duration: AppTheme.durStandard }
                            }

                            // Плёнка наведения
                            Rectangle {
                                anchors.fill: parent
                                radius: parent.radius
                                color: dayMouseArea.pressed
                                       ? AppTheme.statePress
                                       : (dayMouseArea.containsMouse ? AppTheme.stateHover : "transparent")
                                Behavior on color {
                                    ColorAnimation { duration: AppTheme.durMicro }
                                }
                            }

                            Loader {
                                anchors.fill: parent
                                active: isValid && dayInfo.is_holiday
                                sourceComponent: Component { HolidaySparkles {} }
                            }

                            Loader {
                                anchors.fill: parent
                                active: isValid
                                        && !dayInfo.is_holiday
                                        && !dayInfo.is_weekend
                                        && (dayInfo.is_pre_holiday === true)
                                sourceComponent: Component { PreHolidaySparkle {} }
                            }

                            Column {
                                id: rightPanel
                                anchors.top: parent.top
                                anchors.right: parent.right
                                anchors.margins: AppTheme.spaceXS
                                width: AppTheme.spaceL
                                spacing: AppTheme.spaceXS
                                
                                Text { 
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    text: isValid ? dayInfo.day_number : ""
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeSmall
                                    font.weight: AppTheme.weightBold
                                    color: isToday
                                           ? AppTheme.accentBrand
                                           : ((isValid && (dayInfo.is_weekend || dayInfo.is_holiday))
                                              ? AppTheme.accentDanger
                                              : AppTheme.textSecondary)
                                }
                                
                                Rectangle { 
                                    id: compBadgeItem
                                    visible: isValid && dayInfo.has_comp
                                    width: 18
                                    height: 18
                                    radius: AppTheme.radiusPill
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    color: AppTheme.bgTealSoft 
                                    Text {
                                        anchors.centerIn: parent
                                        text: "В"
                                        font.pixelSize: AppTheme.sizeMicro
                                        font.weight: AppTheme.weightBold
                                        color: AppTheme.accentTeal
                                    }
                                }
                                
                                Rectangle {
                                    id: statusBarItem
                                    visible: isValid && dayInfo.status !== ""
                                    width: 18
                                    height: 18
                                    radius: AppTheme.radiusPill
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    color: !isValid
                                           ? "transparent"
                                           : (dayInfo.status === "Б"
                                              ? AppTheme.bgDangerSoft
                                              : (dayInfo.status === "О"
                                                 ? AppTheme.bgWarningSoft
                                                 : AppTheme.bgPurpleSoft))
                                    Text { 
                                        anchors.centerIn: parent
                                        text: isValid ? dayInfo.status : ""
                                        font.pixelSize: AppTheme.sizeMicro
                                        font.weight: AppTheme.weightBold
                                        color: !isValid
                                               ? "transparent"
                                               : (dayInfo.status === "Б"
                                                  ? AppTheme.accentDanger
                                                  : (dayInfo.status === "О"
                                                     ? AppTheme.accentWarning
                                                     : AppTheme.accentPurple))
                                    }
                                }
                            }

                            Flickable {
                                id: dutiesFlickable
                                anchors.top: parent.top
                                anchors.bottom: parent.bottom
                                anchors.left: parent.left
                                anchors.right: rightPanel.left
                                anchors.topMargin: AppTheme.spaceXS
                                anchors.bottomMargin: AppTheme.spaceXS
                                anchors.leftMargin: AppTheme.spaceXS
                                anchors.rightMargin: AppTheme.spaceXS
                                clip: true
                                interactive: false
                                contentHeight: dutyColItem.height
                                
                                WheelHandler {
                                    onWheel: (event) => {
                                        if (dutiesFlickable.contentHeight <= dutiesFlickable.height) return
                                        var newY = dutiesFlickable.contentY - (event.angleDelta.y * 0.3)
                                        var maxY = dutiesFlickable.contentHeight - dutiesFlickable.height
                                        if (newY < 0) newY = 0
                                        if (newY > maxY) newY = maxY
                                        dutiesFlickable.contentY = newY
                                    }
                                }

                                Column {
                                    id: dutyColItem
                                    width: parent.width
                                    spacing: 2
                                    
                                    Repeater {
                                        model: isValid ? dayInfo.duties : null
                                        delegate: Rectangle {
                                            required property var modelData 
                                            property int dutyId: modelData.id !== undefined
                                                                 ? modelData.id
                                                                 : -1
                                            width: parent.width
                                            height: 18
                                            radius: AppTheme.radiusSmall 
                                            color: modelData.is_shift
                                                   ? AppTheme.bgBrandSoft
                                                   : AppTheme.statePress
                                            Text {
                                                anchors.centerIn: parent
                                                text: modelData.text
                                                font.family: AppTheme.fontFamily
                                                font.pixelSize: AppTheme.sizeMicro
                                                font.weight: AppTheme.weightBold
                                                color: modelData.is_shift
                                                       ? AppTheme.accentBrand
                                                       : AppTheme.textPrimary
                                            }
                                        }
                                    }
                                }
                            }

                            Item {
                                id: dayKeyCatcher
                                anchors.fill: parent
                                
                                Keys.onPressed: (event) => {
                                    if (!isValid) return

                                    if (event.key === Qt.Key_Control
                                        || event.key === Qt.Key_Shift
                                        || event.key === Qt.Key_Alt) return

                                    if (event.modifiers & Qt.ControlModifier) {
                                        if (event.key === Qt.Key_C) {
                                            backend.handleClipboard("copy", dayInfo.date_str)
                                            event.accepted = true
                                            return
                                        }
                                        if (event.key === Qt.Key_X) {
                                            mainWindow.explodeAndDelete(
                                                dayInfo.date_str, "all", null,
                                                function() {
                                                    backend.handleClipboard("cut", dayInfo.date_str)
                                                }
                                            )
                                            event.accepted = true
                                            return
                                        }
                                        if (event.key === Qt.Key_V) {
                                            backend.handleClipboard("paste", dayInfo.date_str)
                                            event.accepted = true
                                            return
                                        }
                                    }

                                    if (event.key === Qt.Key_Delete
                                        || event.key === Qt.Key_Backspace) {
                                        mainWindow.explodeAndDelete(
                                            dayInfo.date_str, "all", null,
                                            function() {
                                                backend.handleClipboard("delete", dayInfo.date_str)
                                            }
                                        )
                                        event.accepted = true
                                        return
                                    }

                                    let seq = ""
                                    if (event.modifiers & Qt.ControlModifier) seq += "Ctrl+"
                                    if (event.modifiers & Qt.AltModifier)     seq += "Alt+"
                                    if (event.modifiers & Qt.ShiftModifier)   seq += "Shift+"

                                    let keyText = event.text.toUpperCase()
                                    if (keyText === "") {
                                        if (event.key === Qt.Key_Space) {
                                            keyText = "Space"
                                        } else if (event.key === Qt.Key_Enter
                                                   || event.key === Qt.Key_Return) {
                                            keyText = "Enter"
                                        }
                                    }

                                    if (keyText !== "") {
                                        backend.executeHotkey(seq + keyText, dayInfo.date_str)
                                        event.accepted = true
                                    }
                                }
                            }

                            MouseArea { 
                                id: dayMouseArea
                                anchors.fill: parent
                                enabled: isValid && dayInfo.is_current_month
                                hoverEnabled: true
                                acceptedButtons: Qt.LeftButton | Qt.RightButton
                                
                                onContainsMouseChanged: {
                                    if (containsMouse) dayKeyCatcher.forceActiveFocus()
                                }
                                
                                onClicked: (mouse) => { 
                                    if (!isValid) return
                                    if (mouse.button === Qt.RightButton) {
                                        monthPage.menuDate      = dayInfo.date_str
                                        monthPage.menuIsWeekend = dayInfo.is_weekend
                                        monthPage.menuIsHoliday = dayInfo.is_holiday
                                        monthPage.menuHasDuties = dayInfo.duties.length > 0
                                        monthPage.menuHasComps  = dayInfo.has_comp
                                        monthPage.menuDayCell   = dayCell 
                                        globalContextMenu.popup() 
                                    }
                                }
                                
                                onDoubleClicked: (mouse) => {
                                    if (mouse.button === Qt.LeftButton && isValid) {
                                        dayEventDialog.openForDuty(
                                            dayInfo.date_str, dayCell, mouse.x, mouse.y
                                        )
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // ГЛОБАЛЬНОЕ МЕНЮ КАЛЕНДАРЯ
            AppMenu {
                id: globalContextMenu
                
                background: Item {
                    implicitWidth: 220
                    implicitHeight: globalContextMenu.contentHeight
                    
                    Rectangle { 
                        y: 55
                        width: parent.width
                        height: parent.height - 55
                        color: AppTheme.bgElevated
                        radius: AppTheme.radiusMedium
                        border.color: AppTheme.borderDivider
                        border.width: 1 
                        layer.enabled: true
                        layer.effect: DropShadow {
                            transparentBorder: true
                            color: AppTheme.shadowColor
                            radius: AppTheme.shadowL2Blur
                            verticalOffset: AppTheme.shadowL2Y
                            samples: 25
                        }
                    }
                }
                
                MenuItem {
                    height: 55
                    padding: 0
                    background: Item {} 
                    
                    contentItem: Item {
                        anchors.fill: parent

                        Row {
                            anchors.centerIn: parent
                            spacing: 12
                            
                            Rectangle { 
                                width: 34; height: 34
                                radius: AppTheme.radiusPill
                                color: btnK.pressed
                                       ? AppTheme.statePress
                                       : (btnK.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
                                border.color: AppTheme.borderDivider
                                border.width: 1
                                scale: globalContextMenu.opened ? 1.0 : 0.0
                                Behavior on scale {
                                    NumberAnimation {
                                        duration: AppTheme.durSlow
                                        easing.type: Easing.OutBack
                                    }
                                }
                                Text {
                                    anchors.centerIn: parent
                                    text: "К"
                                    color: AppTheme.accentPurple
                                    font.weight: AppTheme.weightBold
                                    font.pixelSize: AppTheme.sizeBodyLarge
                                } 
                                MouseArea {
                                    id: btnK
                                    anchors.fill: parent
                                    hoverEnabled: true
                                    onClicked: {
                                        backend.setDayStatus(monthPage.menuDate, "К")
                                        globalContextMenu.close()
                                    }
                                }
                                AppToolTip {
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    anchors.bottom: parent.top
                                    anchors.bottomMargin: AppTheme.spaceXS
                                    text: "Командировка"
                                    isVisible: btnK.containsMouse
                                }
                            }
                            
                            Rectangle { 
                                width: 34; height: 34
                                radius: AppTheme.radiusPill
                                color: btnB.pressed
                                       ? AppTheme.statePress
                                       : (btnB.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
                                border.color: AppTheme.borderDivider
                                border.width: 1
                                scale: globalContextMenu.opened ? 1.0 : 0.0
                                Behavior on scale {
                                    SequentialAnimation {
                                        PauseAnimation { duration: 40 }
                                        NumberAnimation {
                                            duration: AppTheme.durSlow
                                            easing.type: Easing.OutBack
                                        }
                                    }
                                }
                                Text {
                                    anchors.centerIn: parent
                                    text: "Б"
                                    color: AppTheme.accentDanger
                                    font.weight: AppTheme.weightBold
                                    font.pixelSize: AppTheme.sizeBodyLarge
                                } 
                                MouseArea {
                                    id: btnB
                                    anchors.fill: parent
                                    hoverEnabled: true
                                    onClicked: {
                                        backend.setDayStatus(monthPage.menuDate, "Б")
                                        globalContextMenu.close()
                                    }
                                }
                                AppToolTip {
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    anchors.bottom: parent.top
                                    anchors.bottomMargin: AppTheme.spaceXS
                                    text: "Больничный"
                                    isVisible: btnB.containsMouse
                                }
                            }
                            
                            Rectangle { 
                                width: 34; height: 34
                                radius: AppTheme.radiusPill
                                color: btnO.pressed
                                       ? AppTheme.statePress
                                       : (btnO.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
                                border.color: AppTheme.borderDivider
                                border.width: 1
                                scale: globalContextMenu.opened ? 1.0 : 0.0
                                Behavior on scale {
                                    SequentialAnimation {
                                        PauseAnimation { duration: 80 }
                                        NumberAnimation {
                                            duration: AppTheme.durSlow
                                            easing.type: Easing.OutBack
                                        }
                                    }
                                }
                                Text {
                                    anchors.centerIn: parent
                                    text: "О"
                                    color: AppTheme.accentWarning
                                    font.weight: AppTheme.weightBold
                                    font.pixelSize: AppTheme.sizeBodyLarge
                                } 
                                MouseArea {
                                    id: btnO
                                    anchors.fill: parent
                                    hoverEnabled: true
                                    onClicked: {
                                        backend.setDayStatus(monthPage.menuDate, "О")
                                        globalContextMenu.close()
                                    }
                                }
                                AppToolTip {
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    anchors.bottom: parent.top
                                    anchors.bottomMargin: AppTheme.spaceXS
                                    text: "Отпуск"
                                    isVisible: btnO.containsMouse
                                }
                            }
                            
                            Rectangle { 
                                width: 34; height: 34
                                radius: AppTheme.radiusPill
                                color: btnClear.pressed
                                       ? AppTheme.statePress
                                       : (btnClear.containsMouse ? AppTheme.stateHover : AppTheme.bgSurface)
                                border.color: AppTheme.borderDivider
                                border.width: 1
                                scale: globalContextMenu.opened ? 1.0 : 0.0
                                Behavior on scale {
                                    SequentialAnimation {
                                        PauseAnimation { duration: 120 }
                                        NumberAnimation {
                                            duration: AppTheme.durSlow
                                            easing.type: Easing.OutBack
                                        }
                                    }
                                }
                                IconImage {
                                    anchors.centerIn: parent
                                    source: "../icons/trash.svg"
                                    width: AppTheme.iconMedium
                                    height: AppTheme.iconMedium
                                    color: btnClear.containsMouse
                                           ? AppTheme.accentDanger
                                           : AppTheme.textSecondary
                                } 
                                MouseArea {
                                    id: btnClear
                                    anchors.fill: parent
                                    hoverEnabled: true
                                    onClicked: {
                                        mainWindow.explodeAndDelete(
                                            monthPage.menuDate, "status", null,
                                            function() {
                                                backend.setDayStatus(monthPage.menuDate, "")
                                            }
                                        )
                                        globalContextMenu.close()
                                    }
                                }
                                AppToolTip {
                                    anchors.horizontalCenter: parent.horizontalCenter
                                    anchors.bottom: parent.top
                                    anchors.bottomMargin: AppTheme.spaceXS
                                    text: "Удалить статус"
                                    isVisible: btnClear.containsMouse
                                }
                            }
                        }
                    }
                }
                
                AppMenuItem {
                    text: "Открыть день"
                    iconSource: "../icons/edit.svg"
                    onClicked: {
                        backend.loadDayDetails(monthPage.menuDate)
                        dayInspector.targetDate = monthPage.menuDate
                        dayInspector.show()
                        globalContextMenu.close()
                    }
                }

                AppMenuSeparator {}
                
                AppMenuItem {
                    text: "Добавить дежурство"
                    iconSource: "../icons/clock.svg"
                    onClicked: {
                        dayEventDialog.openForDuty(
                            monthPage.menuDate, globalContextMenu.parent, 0, 0
                        )
                        globalContextMenu.close()
                    }
                    Rectangle { 
                        visible: monthPage.menuHasDuties
                        width: 26; height: 26
                        radius: AppTheme.radiusSmall
                        anchors.verticalCenter: parent.verticalCenter
                        anchors.right: parent.right
                        anchors.rightMargin: AppTheme.spaceS
                        color: dutyDelMouse.containsMouse ? AppTheme.bgDangerSoft : "transparent"
                        IconImage {
                            anchors.centerIn: parent
                            source: "../icons/trash.svg"
                            width: AppTheme.iconSmall
                            height: AppTheme.iconSmall
                            color: dutyDelMouse.containsMouse
                                   ? AppTheme.accentDanger
                                   : AppTheme.textTertiary
                        }
                        MouseArea {
                            id: dutyDelMouse
                            anchors.fill: parent
                            hoverEnabled: true
                            onClicked: {
                                mainWindow.explodeAndDelete(
                                    monthPage.menuDate, "duty", null,
                                    function() { backend.clearDayDuties(monthPage.menuDate) }
                                )
                                globalContextMenu.close()
                            }
                        } 
                    } 
                }
                
                AppMenuItem {
                    text: "Добавить компенсацию"
                    iconSource: "../icons/rest.svg"
                    onClicked: {
                        dayEventDialog.openForComp(
                            monthPage.menuDate, globalContextMenu.parent, 0, 0
                        )
                        globalContextMenu.close()
                    }
                    Rectangle { 
                        visible: monthPage.menuHasComps
                        width: 26; height: 26
                        radius: AppTheme.radiusSmall
                        anchors.verticalCenter: parent.verticalCenter
                        anchors.right: parent.right
                        anchors.rightMargin: AppTheme.spaceS
                        color: compDelMouse.containsMouse ? AppTheme.bgDangerSoft : "transparent"
                        IconImage {
                            anchors.centerIn: parent
                            source: "../icons/trash.svg"
                            width: AppTheme.iconSmall
                            height: AppTheme.iconSmall
                            color: compDelMouse.containsMouse
                                   ? AppTheme.accentDanger
                                   : AppTheme.textTertiary
                        }
                        MouseArea {
                            id: compDelMouse
                            anchors.fill: parent
                            hoverEnabled: true
                            onClicked: {
                                mainWindow.explodeAndDelete(
                                    monthPage.menuDate, "comp", null,
                                    function() { backend.clearDayCompensations(monthPage.menuDate) }
                                )
                                globalContextMenu.close()
                            }
                        } 
                    } 
                }
                
                AppMenuSeparator {}

                AppMenuItem {
                    visible: monthPage.menuIsWeekend || monthPage.menuIsHoliday
                    text: "Сделать рабочим"
                    customColor: AppTheme.accentTeal
                    onClicked: {
                        backend.setDayType(monthPage.menuDate, "work")
                        globalContextMenu.close()
                    }
                }
                AppMenuItem {
                    visible: !monthPage.menuIsWeekend
                    text: "Сделать выходным"
                    customColor: AppTheme.accentDanger
                    onClicked: {
                        backend.setDayType(monthPage.menuDate, "weekend")
                        globalContextMenu.close()
                    }
                }
                AppMenuItem {
                    visible: !monthPage.menuIsHoliday
                    text: "Сделать праздничным"
                    customColor: AppTheme.accentPurple
                    onClicked: {
                        backend.setDayType(monthPage.menuDate, "holiday")
                        globalContextMenu.close()
                    }
                }
            }                
        }

        // ==========================================
        // СТРАНИЦА 1: ГОДОВАЯ ПАНОРАМА
        // ==========================================
        Item {
            Rectangle {
                anchors.top: parent.top
                anchors.bottom: parent.bottom
                anchors.left: parent.left
                anchors.right: parent.right
                anchors.margins: AppTheme.spaceL
                color: "transparent"
                
                Text { 
                    visible: backend.yearlyData.length === 0
                    anchors.centerIn: parent
                    text: "Выберите сотрудника слева"
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeH3 
                }
                
                Column {
                    id: monthLabels
                    visible: backend.yearlyData.length > 0
                    anchors.left: parent.left
                    anchors.top: yearGrid.top
                    anchors.bottom: yearGrid.bottom
                    width: 40
                    
                    Repeater { 
                        model: ["Янв", "Фев", "Мар", "Апр", "Май", "Июн",
                                "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
                        Text { 
                            height: yearGrid.height / 12
                            text: modelData
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                            verticalAlignment: Text.AlignVCenter 
                        } 
                    }
                }
                
                Row {
                    id: daysScale
                    visible: backend.yearlyData.length > 0
                    anchors.left: yearGrid.left
                    anchors.right: yearGrid.right
                    anchors.top: parent.top
                    height: 24
                    
                    Repeater { 
                        model: 31
                        Item { 
                            width: daysScale.width / 31
                            height: parent.height
                            Text { 
                                anchors.centerIn: parent
                                visible: index === 0  || index === 4  || index === 9
                                      || index === 14 || index === 19 || index === 24
                                      || index === 30
                                text: index + 1
                                color: AppTheme.textSecondary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                font.weight: AppTheme.weightBold 
                            } 
                        } 
                    }
                }
                
                ListView {
                    id: yearGrid
                    visible: backend.yearlyData.length > 0
                    anchors.left: monthLabels.right
                    anchors.top: daysScale.bottom
                    anchors.topMargin: AppTheme.spaceS
                    anchors.right: parent.right
                    height: (width / 31) * 12
                    interactive: false
                    reuseItems: true 
                    model: backend.yearlyData
                    
                    delegate: Row {
                        required property var modelData
                        property var rowInfo: modelData
                        width: ListView.view.width
                        height: ListView.view.height / 12
                        spacing: 0
                        
                        Repeater {
                            model: rowInfo
                            
                            delegate: Item {
                                required property var modelData
                                property var cellInfo: modelData
                                width: parent.width / 31
                                height: parent.height
                                
                                Rectangle {
                                    anchors.fill: parent
                                    anchors.margins: 2
                                    radius: AppTheme.radiusSmall
                                    visible: cellInfo.is_real
                                    color: !cellInfo
                                           ? "transparent"
                                           : (cellInfo.type === "duty"
                                              ? (cellInfo.val === "1"
                                                 ? AppTheme.accentBrand
                                                 : AppTheme.bgBrandSoft)
                                              : (cellInfo.is_weekend
                                                 ? AppTheme.bgDangerSoft
                                                 : AppTheme.bgCell))
                                    
                                    Text { 
                                        anchors.centerIn: parent
                                        font.family: AppTheme.fontFamily
                                        font.pixelSize: AppTheme.sizeSmall
                                        font.weight: AppTheme.weightBold
                                        text: !cellInfo
                                              ? ""
                                              : (cellInfo.type === "status"
                                                 ? cellInfo.val
                                                 : (cellInfo.type === "comp" ? "В" : ""))
                                        color: !cellInfo
                                               ? "transparent"
                                               : (cellInfo.type === "status"
                                                  ? (cellInfo.val === "Б"
                                                     ? AppTheme.accentDanger
                                                     : (cellInfo.val === "О"
                                                        ? AppTheme.accentWarning
                                                        : AppTheme.accentPurple))
                                                  : (cellInfo.type === "comp"
                                                     ? AppTheme.accentTeal
                                                     : "transparent"))
                                    }
                                    
                                    MouseArea { 
                                        id: yearMouseArea
                                        anchors.fill: parent
                                        hoverEnabled: true
                                        
                                        Rectangle { 
                                            anchors.fill: parent
                                            color: AppTheme.stateHover
                                            opacity: parent.containsMouse ? 1.0 : 0.0
                                            radius: AppTheme.radiusSmall
                                            Behavior on opacity {
                                                NumberAnimation { duration: AppTheme.durMicro }
                                            } 
                                        } 
                                        
                                        onClicked: { 
                                            backend.jumpToMonth(cellInfo.month)
                                            root.isYearView = false
                                            calendarStack.currentIndex = 0 
                                        } 
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
