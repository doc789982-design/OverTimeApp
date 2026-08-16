import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

TextField {
    id: root

    property string label: ""
    property bool isRequired: false 
    property string selectedDate: "" 
    property color cutoutColor: AppTheme.bgModal
    
    property bool isFloated: root.text.length > 0 || root.activeFocus || calendarDialog.opened

    property int currentMonth: new Date().getMonth()
    property int currentYear: new Date().getFullYear()
    property int targetMonth: currentMonth
    property int targetYear: currentYear

    implicitHeight: 44 
    Layout.fillWidth: true
    
    leftPadding: AppTheme.spaceM
    rightPadding: 36 // Под иконку
    verticalAlignment: TextInput.AlignVCenter
    
    color: root.enabled ? AppTheme.textPrimary : AppTheme.textDisabled
    font.family: AppTheme.fontFamily
    font.pixelSize: AppTheme.sizeBody
    
    placeholderText: "ДД.ММ.ГГГГ"
    placeholderTextColor: {
        if ((root.activeFocus || calendarDialog.opened) && floatingLabel.y < 0) {
            return AppTheme.textTertiary; 
        } else {
            return Qt.rgba(AppTheme.textTertiary.r, AppTheme.textTertiary.g, AppTheme.textTertiary.b, 0.0);
        }
    }
    Behavior on placeholderTextColor { ColorAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutQuad } }
    
    focusPolicy: Qt.StrongFocus
    cursorDelegate: AppCursorDelegate {}

    onActiveFocusChanged: { if (activeFocus) Qt.callLater(function() { root.selectAll() }) }

    onTextEdited: {
        let raw = text.replace(/[^0-9]/g, '')
        if (raw.length > 8) raw = raw.substring(0, 8)
        let formatted = ''
        if (raw.length > 0) formatted += raw.substring(0, 2)
        if (raw.length >= 3) formatted += '.' + raw.substring(2, 4)
        if (raw.length >= 5) formatted += '.' + raw.substring(4, 8)
        if (text !== formatted) { text = formatted; cursorPosition = formatted.length }
    }

    onEditingFinished: root.parseManualInput(root.text)

    // ==========================================
    // РАМКА
    // ==========================================
    background: Rectangle {
        color: "transparent"
        radius: AppTheme.radiusMedium
        
        border.color: !root.enabled ? AppTheme.borderDisabled :
                      (root.activeFocus || calendarDialog.opened ? AppTheme.borderFocus : 
                      (root.hovered ? AppTheme.textSecondary : AppTheme.borderInput))
        
        border.width: (root.activeFocus || calendarDialog.opened) ? AppTheme.focusWidth : 1
        Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ИКОНКА КАЛЕНДАРЯ
    Item {
        width: 36; height: 36
        anchors.right: parent.right
        anchors.verticalCenter: parent.verticalCenter
        
        IconImage { 
            anchors.centerIn: parent
            source: "../icons/calendar.svg"
            width: AppTheme.iconMedium; height: AppTheme.iconMedium
            color: calendarDialog.opened ? AppTheme.accentBrand : AppTheme.textSecondary
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } } 
        }
        MouseArea {
            anchors.fill: parent; cursorShape: Qt.PointingHandCursor
            onClicked: { 
                if (calendarDialog.opened) calendarDialog.close()
                else calendarDialog.showAt(root, root.width / 2, root.height) 
            }
        }
    }

    // ==========================================
    // ИДЕАЛЬНЫЙ ЛАСТИК И ЛЕЙБЛ
    // ==========================================
    Rectangle {
        color: root.cutoutColor
        x: floatingLabel.x - 4
        y: -2 
        height: 4 
        width: (floatingLabel.width * floatingLabel.scale) + 8
        opacity: root.isFloated ? 1.0 : 0.0
        Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
    }

    Row {
        id: floatingLabel
        x: AppTheme.spaceS
        y: root.isFloated ? -(height * 0.75) / 2 : (root.height - height) / 2
        
        scale: root.isFloated ? 0.75 : 1.0
        transformOrigin: Item.TopLeft
        
        Behavior on y { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }
        Behavior on scale { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }

        spacing: AppTheme.spaceMicro

        Text {
            text: root.label
            color: !root.enabled ? AppTheme.textDisabled : 
                   ((root.activeFocus || calendarDialog.opened) ? AppTheme.accentBrand : AppTheme.textSecondary)
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
            font.weight: root.isFloated ? AppTheme.weightMedium : AppTheme.weightRegular
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        }
        
        Text {
            visible: root.isRequired
            text: "*"
            color: root.enabled ? AppTheme.accentDanger : AppTheme.textDisabled 
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
        }
    }

    // ==========================================
    // 3. КАЛЕНДАРЬ (Выпадающее окно)
    // ==========================================
    AppDialog {
        id: calendarDialog
        parent: Overlay.overlay 
        width: 330 
        title: ""
        
        showFooter: false
        
        onOpened: root.syncCalendarToDate(root.selectedDate)

        // ШАПКА КАЛЕНДАРЯ
        RowLayout {
            width: parent.width
            
            Rectangle {
                width: 28; height: 28; radius: AppTheme.radiusSmall
                color: prevHover.pressed ? AppTheme.statePress : (prevHover.containsMouse ? AppTheme.stateHover : "transparent")
                Text { anchors.centerIn: parent; text: "‹"; color: AppTheme.textSecondary; font.pixelSize: AppTheme.sizeH2 }
                MouseArea { id: prevHover; anchors.fill: parent; hoverEnabled: true; onClicked: root.changeMonth(-1) }
            }
            
            Text {
                Layout.fillWidth: true; horizontalAlignment: Text.AlignHCenter
                text: root.getMonthName(root.currentMonth) + " " + root.currentYear
                color: AppTheme.textPrimary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; font.weight: AppTheme.weightBold
            }
            
            Rectangle {
                width: 28; height: 28; radius: AppTheme.radiusSmall
                color: nextHover.pressed ? AppTheme.statePress : (nextHover.containsMouse ? AppTheme.stateHover : "transparent")
                Text { anchors.centerIn: parent; text: "›"; color: AppTheme.textSecondary; font.pixelSize: AppTheme.sizeH2 }
                MouseArea { id: nextHover; anchors.fill: parent; hoverEnabled: true; onClicked: root.changeMonth(1) }
            }
        }

        // ДНИ НЕДЕЛИ
        Row {
            width: parent.width
            Repeater {
                model: ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
                Item {
                    width: parent.width / 7
                    height: 24
                    Text { 
                        anchors.centerIn: parent
                        text: modelData
                        color: index >= 5 ? AppTheme.accentDanger : AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        font.weight: AppTheme.weightMedium 
                    }
                }
            }
        }

        // СЕТКА ДНЕЙ
        Item {
            id: gridContainer
            width: parent.width
            height: daysGrid.implicitHeight

            GridLayout {
                id: daysGrid
                anchors.fill: parent
                columns: 7
                rowSpacing: AppTheme.spaceXXS
                columnSpacing: 0
                
                Repeater {
                    id: daysRepeater
                    model: 42 
                    
                    Rectangle {
                        width: gridContainer.width / 7
                        height: width
                        radius: AppTheme.radiusPill // Круглые ячейки
                        
                        property var dayInfo: root.getDayInfo(index)
                        property bool isSelected: root.selectedDate === dayInfo.isoDate
                        property bool isToday: root.getTodayIso() === dayInfo.isoDate
                        
                        color: isSelected ? AppTheme.accentBrand : 
                               (isToday ? AppTheme.bgBrandSoft : 
                               (dayHover.pressed ? AppTheme.statePress : 
                               (dayHover.containsMouse ? AppTheme.stateHover : "transparent")))
                               
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        
                        Text {
                            anchors.centerIn: parent
                            text: dayInfo.day > 0 ? dayInfo.day : ""
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: (isSelected || isToday) ? AppTheme.weightBold : AppTheme.weightRegular
                            
                            color: { 
                                if (isSelected) return AppTheme.textOnAccent
                                if (!dayInfo.isCurrentMonth) return AppTheme.textTertiary
                                if (isToday) return AppTheme.accentBrand
                                return AppTheme.textPrimary 
                            }
                        }
                        
                        MouseArea {
                            id: dayHover
                            anchors.fill: parent
                            hoverEnabled: dayInfo.day > 0
                            visible: dayInfo.day > 0
                            cursorShape: Qt.PointingHandCursor
                            
                            onClicked: {
                                root.selectedDate = dayInfo.isoDate
                                calendarDialog.close()
                            }
                        }
                    }
                }
            }
        }
    }

    SequentialAnimation {
        id: monthTransitionAnim
        
        ParallelAnimation { 
            NumberAnimation { target: gridContainer; property: "opacity"; to: 0.0; duration: AppTheme.durMicro }
            NumberAnimation { target: gridContainer; property: "scale"; to: 0.95; duration: AppTheme.durMicro } 
        }
        
        ScriptAction { 
            script: { 
                root.currentMonth = root.targetMonth
                root.currentYear = root.targetYear
                daysRepeater.model = 0
                daysRepeater.model = 42 
            } 
        }
        
        ParallelAnimation { 
            NumberAnimation { target: gridContainer; property: "opacity"; to: 1.0; duration: AppTheme.durFast }
            NumberAnimation { target: gridContainer; property: "scale"; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } 
        }
    }

    onSelectedDateChanged: {
        if (selectedDate === "") { root.text = ""; return }
        let parts = selectedDate.split("-")
        if (parts.length === 3) root.text = parts[2] + "." + parts[1] + "." + parts[0]
    }

    function parseManualInput(txt) {
        let parts = txt.split(".")
        if (parts.length === 3 && parts[2].length === 4) {
            let d = parseInt(parts[0]); let m = parseInt(parts[1]); let y = parseInt(parts[2])
            if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
                let monthStr = ("0" + m).slice(-2); let dayStr = ("0" + d).slice(-2)
                root.selectedDate = y + "-" + monthStr + "-" + dayStr
                return
            }
        }
        root.selectedDateChanged() 
    }

    function getTodayIso() { 
        let d = new Date()
        return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2) 
    }

    function syncCalendarToDate(isoStr) {
        if (isoStr !== "") { 
            let parts = isoStr.split("-")
            if (parts.length === 3) { currentYear = parseInt(parts[0]); currentMonth = parseInt(parts[1]) - 1 }
        } else { 
            let d = new Date()
            currentYear = d.getFullYear(); currentMonth = d.getMonth() 
        }
        daysRepeater.model = 0; daysRepeater.model = 42
    }

    function changeMonth(delta) {
        let m = currentMonth + delta; let y = currentYear
        if (m > 11) { m = 0; y++ }; if (m < 0) { m = 11; y-- }
        targetMonth = m; targetYear = y
        monthTransitionAnim.restart() 
    }

    function getMonthName(m) { 
        return ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"][m] 
    }

    function getDayInfo(index) {
        let firstDay = new Date(currentYear, currentMonth, 1)
        let startDay = firstDay.getDay()
        if (startDay === 0) startDay = 7
        startDay -= 1
        
        let daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate()
        let daysInPrev = new Date(currentYear, currentMonth, 0).getDate()
        
        let dNum = 0; let isCur = false; let iso = ""
        
        if (index < startDay) { 
            dNum = daysInPrev - (startDay - index - 1)
            let pM = currentMonth === 0 ? 12 : currentMonth; let pY = currentMonth === 0 ? currentYear - 1 : currentYear
            iso = pY + "-" + ("0" + pM).slice(-2) + "-" + ("0" + dNum).slice(-2)
        } else if (index >= startDay && index < startDay + daysInMonth) { 
            dNum = index - startDay + 1; isCur = true
            iso = currentYear + "-" + ("0" + (currentMonth + 1)).slice(-2) + "-" + ("0" + dNum).slice(-2)
        } else { 
            dNum = index - (startDay + daysInMonth) + 1
            let nM = currentMonth === 11 ? 1 : currentMonth + 2; let nY = currentMonth === 11 ? currentYear + 1 : currentYear
            iso = nY + "-" + ("0" + nM).slice(-2) + "-" + ("0" + dNum).slice(-2) 
        }
        
        return { day: dNum, isCurrentMonth: isCur, isoDate: iso }
    }
}