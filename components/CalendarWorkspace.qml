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
    // Красивая дата для сообщений: "2026-08-24" -> "24.08.2026"
    function fmtRuDate(iso) {
        if (!iso) return ""
        let p = iso.split("-")
        if (p.length !== 3) return iso
        return p[2] + "." + p[1] + "." + p[0]
    }

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
                            color: {
                                if (index === 6)
                                    return AppTheme.accentDanger
                                if (backend.isSelectedEmployeeShiftedWeekends)
                                    return index === 0 ? AppTheme.accentDanger : AppTheme.textSecondary
                                return index >= 5 ? AppTheme.accentDanger : AppTheme.textSecondary
                            }
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightMedium 
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
                                // ОПТИМИЗАЦИЯ: искры только в видимых ячейках текущего
                                // месяца (раньше крутились и в прозрачных ячейках соседних)
                                active: isValid && dayInfo.is_holiday && dayInfo.is_current_month
                                sourceComponent: Component { HolidaySparkles {} }
                            }

                            Loader {
                                anchors.fill: parent
                                active: isValid
                                        && dayInfo.is_current_month
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
                                
                                // ЛКМ и ПКМ — открываем меню дня (как контекстное меню).
                                onClicked: (mouse) => { 
                                    if (!isValid) return
                                    dayMenu.openFromCell(
                                        dayCell,
                                        dayInfo.date_str,
                                        dayInfo.is_weekend,
                                        dayInfo.is_holiday,
                                        dayInfo.duties.length > 0,
                                        dayInfo.has_comp
                                    )
                                }
                            }
                        }
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
                                    color: {
                                        if (!cellInfo)
                                            return "transparent"
                                        if (cellInfo.type === "duty")
                                            return cellInfo.val === "1"
                                                   ? AppTheme.yearDutyShift
                                                   : AppTheme.yearDutyExtra
                                        if (cellInfo.is_weekend)
                                            return AppTheme.yearWeekend
                                        return AppTheme.bgCell
                                    }
                                    
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

    // ═══════════════════════════════════════════════════════════════
    // ПУСТОЕ СОСТОЯНИЕ (СОТРУДНИК НЕ ВЫБРАН)
    // Единая накладка поверх обеих вкладок (Месяц и Год). Живёт вне
    // StackLayout, поэтому при переключении вкладок анимация маскота
    // не перезапускается — она идёт непрерывно, пока нет сотрудника.
    // Когда сотрудник выбран, накладка скрывается и открывается
    // календарь или матрица года.
    // ═══════════════════════════════════════════════════════════════
    Item {
        id: emptyStateOverlay
        anchors.top: calendarHeader.bottom
        anchors.bottom: unifiedSummaryPanel.top
        anchors.bottomMargin: AppTheme.spaceM
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
}
