// ========================================
// НАЧАЛО ФАЙЛА: AppTumblerTime.qml
// ========================================

import QtQuick
import QtQuick.Controls

Item {
    id: root
    
    implicitWidth: 100
    implicitHeight: 36

    property int hours: 8
    property int minutes: 0

    function pad(num) {
        return num < 10 ? "0" + num : num.toString()
    }

    Row {
        anchors.centerIn: parent
        spacing: AppTheme.spaceXXS

        // --- ПОЛЕ ЧАСОВ ---
        TextField {
            id: hrInput
            width: 42
            height: 36
            
            // Убираем стандартные отступы Qt, чтобы текст встал ровно по центру
            topPadding: 0
            bottomPadding: 0
            
            text: root.pad(root.hours)
            horizontalAlignment: TextInput.AlignHCenter
            verticalAlignment: TextInput.AlignVCenter
            
            color: AppTheme.textPrimary
            font.family: AppTheme.fontFamily
            // ИСПРАВЛЕНО: Уменьшили шрифт с sizeH3 (28px) до sizeH5 (20px)
            font.pixelSize: AppTheme.sizeH5 
            font.weight: AppTheme.weightBold

            validator: RegularExpressionValidator { regularExpression: /^[0-9]{1,2}$/ }
            focusPolicy: Qt.StrongFocus
            cursorDelegate: AppCursorDelegate {}

            background: Rectangle {
                color: AppTheme.bgInput
                radius: AppTheme.radiusMedium
                
                border.color: hrInput.activeFocus ? "transparent" : (hrInput.hovered ? AppTheme.textSecondary : AppTheme.borderInput)
                border.width: 1
                
                Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }

                // Кольцо фокуса
                Rectangle {
                    anchors.fill: parent
                    radius: parent.radius
                    color: "transparent"
                    border.color: AppTheme.borderFocus
                    border.width: AppTheme.focusWidth
                    opacity: hrInput.activeFocus ? 1.0 : 0.0
                    Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }
                }
            }

            onEditingFinished: {
                let val = parseInt(text)
                if (isNaN(val)) val = 0
                if (val > 23) val = 23 
                root.hours = val
                text = root.pad(root.hours)
                hrInput.focus = false
            }

            onActiveFocusChanged: { if (activeFocus) selectAll() }

            WheelHandler {
                onWheel: (event) => {
                    hrInput.focus = false 
                    let val = root.hours
                    if (event.angleDelta.y > 0) val = (val + 1) % 24
                    else val = (val - 1 + 24) % 24
                    root.hours = val
                    hrInput.text = root.pad(root.hours)
                }
            }
        }

        // --- РАЗДЕЛИТЕЛЬ ---
        Text {
            anchors.verticalCenter: parent.verticalCenter
            text: ":"
            color: AppTheme.textSecondary
            font.family: AppTheme.fontFamily
            // ИСПРАВЛЕНО: Уменьшили размер двоеточия с sizeH2 (32px) до sizeH4 (24px)
            font.pixelSize: AppTheme.sizeH4 
            font.weight: AppTheme.weightBold
            anchors.verticalCenterOffset: -2 // Чуть-чуть приподнимаем двоеточие визуально
        }

        // --- ПОЛЕ МИНУТ ---
        TextField {
            id: minInput
            width: 42
            height: 36
            
            topPadding: 0
            bottomPadding: 0
            
            text: root.pad(root.minutes)
            horizontalAlignment: TextInput.AlignHCenter
            verticalAlignment: TextInput.AlignVCenter
            
            color: AppTheme.textPrimary
            font.family: AppTheme.fontFamily
            // ИСПРАВЛЕНО: Уменьшили шрифт с sizeH3 (28px) до sizeH5 (20px)
            font.pixelSize: AppTheme.sizeH5 
            font.weight: AppTheme.weightBold

            validator: RegularExpressionValidator { regularExpression: /^[0-9]{1,2}$/ }
            focusPolicy: Qt.StrongFocus
            cursorDelegate: AppCursorDelegate {}

            background: Rectangle {
                color: AppTheme.bgInput
                radius: AppTheme.radiusMedium
                
                border.color: minInput.activeFocus ? "transparent" : (minInput.hovered ? AppTheme.textSecondary : AppTheme.borderInput)
                border.width: 1
                
                Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }

                // Кольцо фокуса
                Rectangle {
                    anchors.fill: parent
                    radius: parent.radius
                    color: "transparent"
                    border.color: AppTheme.borderFocus
                    border.width: AppTheme.focusWidth
                    opacity: minInput.activeFocus ? 1.0 : 0.0
                    Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }
                }
            }

            onEditingFinished: {
                let val = parseInt(text)
                if (isNaN(val)) val = 0
                if (val > 59) val = 59 
                root.minutes = val
                text = root.pad(root.minutes)
                minInput.focus = false
            }

            onActiveFocusChanged: { if (activeFocus) selectAll() }

            WheelHandler {
                onWheel: (event) => {
                    minInput.focus = false
                    let val = root.minutes
                    
                    // Крутим колесико ВВЕРХ - прибавляем 1 минуту
                    if (event.angleDelta.y > 0) {
                        val = (val + 1) % 60 
                    } 
                    // Крутим колесико ВНИЗ - отнимаем 1 минуту
                    else {
                        val = (val - 1 + 60) % 60
                    }
                    
                    root.minutes = val
                    minInput.text = root.pad(root.minutes)
                }
            }
        }
    }

    onHoursChanged: { if (!hrInput.activeFocus) hrInput.text = root.pad(hours) }
    onMinutesChanged: { if (!minInput.activeFocus) minInput.text = root.pad(minutes) }
}