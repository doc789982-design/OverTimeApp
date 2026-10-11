import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

// ============================================================
// ПОЛЕ ДАТЫ: базовая механика — в AppFieldBase, здесь только
// дата: маска ввода ДД.ММ.ГГГГ, иконка календаря и выпадающий
// календарь. Пока календарь открыт, поле «горит» (fieldActive).
// ============================================================
AppFieldBase {
    id: root

    property string selectedDate: ""

    // Открытый календарь — тоже «активность» поля: рамка и метка горят
    fieldActive: root.activeFocus || calendarDialog.opened

    rightPadding: 36 // Под иконку
    placeholderText: "ДД.ММ.ГГГГ"

    property int currentMonth: new Date().getMonth()
    property int currentYear: new Date().getFullYear()
    property int targetMonth: currentMonth
    property int targetYear: currentYear

    onTextEdited: {
        let raw = text.replace(/[^0-9]/g, '')
        if (raw.length > 8) raw = raw.substring(0, 8)
        let formatted = ''
        if (raw.length > 0) formatted += raw.substring(0, 2)
        if (raw.length >= 3) formatted += '.' + raw.substring(2, 4)
        if (raw.length >= 5) formatted += '.' + raw.substring(4, 8)
        if (text !== formatted) { text = formatted; cursorPosition = formatted.length }
        if (root.hasError) root.hasError = false   // ввод гасит вспышку ошибки
    }

    onEditingFinished: root.parseManualInput(root.text)

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
        HoverHandler { cursorShape: Qt.PointingHandCursor }
        MouseArea {
            anchors.fill: parent; cursorShape: Qt.PointingHandCursor
            onClicked: {
                if (calendarDialog.opened) calendarDialog.close()
                else calendarDialog.showAt(root, root.width / 2, root.height)
            }
        }
    }

    // ==========================================
    // КАЛЕНДАРЬ (Выпадающее окно)
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
                width: 32; height: 32; radius: AppTheme.radiusSmall
                color: prevHover.pressed ? AppTheme.statePress : (prevHover.containsMouse ? AppTheme.stateHover : "transparent")
                Text { anchors.centerIn: parent; text: "‹"; color: AppTheme.textSecondary; font.pixelSize: AppTheme.sizeH2 }
                HoverHandler { cursorShape: Qt.PointingHandCursor }
                MouseArea { id: prevHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.changeMonth(-1) }
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
                HoverHandler { cursorShape: Qt.PointingHandCursor }
                MouseArea { id: nextHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.changeMonth(1) }
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

                        HoverHandler { cursorShape: Qt.PointingHandCursor }
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
        if (root.hasError) root.hasError = false   // ввод даты гасит вспышку ошибки
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
