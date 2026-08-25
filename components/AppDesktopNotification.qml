import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Window 
import Qt5Compat.GraphicalEffects

Window {
    id: root
    
    // Системное окно поверх всего Windows
    flags: Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
    color: "transparent"

    width: 400
    // Высота окна = высота плашки + 32 пикселя на тени (без зацикливания!)
    height: bgRect.height + 32 

    property string message: "Не забудьте сделать табель, ознакомить с ним всех сотрудников и сдать его в кадровое подразделение до 5-го числа месяца. Скрыть до следующего напоминания."
    property string actionText: ""   // Если задан — показываем кнопку действия

    signal actionTriggered()

    Rectangle {
        id: bgRect
        // Жестко задаем координаты вместо anchors.fill, чтобы не было зацикливания
        x: 16
        y: 16
        width: parent.width - 32
        height: contentLayout.implicitHeight + 32

        color: AppTheme.bgElevated
        radius: AppTheme.radiusMedium
        border.color: AppTheme.borderDivider
        border.width: 1

        // Тень-картинка вместо вычисляемой (Level 5, сдвиг как L3)
        AppShadow { level: 5; yOffset: 4 }

        RowLayout {
            id: contentLayout
            anchors.fill: parent
            anchors.leftMargin: AppTheme.spaceL
            anchors.rightMargin: AppTheme.spaceM
            spacing: AppTheme.spaceL

            Text {
                Layout.fillWidth: true
                Layout.alignment: Qt.AlignVCenter
                text: root.message
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                wrapMode: Text.WordWrap
                lineHeight: 1.3
            }

            AppButton {
                Layout.alignment: Qt.AlignVCenter
                text: "Ок"
                variant: "secondary" 
                implicitHeight: 32
                onClicked: closeAnim.start() 
            }

            AppButton {
                Layout.alignment: Qt.AlignVCenter
                visible: root.actionText !== ""
                text: root.actionText
                variant: "primary"
                implicitHeight: 32
                onClicked: { root.actionTriggered(); closeAnim.start() }
            }
        }
    }

    ParallelAnimation {
        id: openAnim
        NumberAnimation { id: openXAnim; target: root; property: "x"; duration: 500; easing.type: Easing.OutBack }
        NumberAnimation { target: root; property: "opacity"; from: 0.0; to: 1.0; duration: 300 }
    }

    ParallelAnimation {
        id: closeAnim
        NumberAnimation { id: closeXAnim; target: root; property: "x"; duration: 400; easing.type: Easing.InBack }
        NumberAnimation { target: root; property: "opacity"; from: 1.0; to: 0.0; duration: 300 }
        onFinished: root.hide() 
    }

    function showNotification() {
        // Обычное напоминание: без кнопки действия
        root.actionText = ""
        _show()
    }

    function showCustom(msg, action) {
        // Уведомление с текстом и (необязательно) кнопкой действия
        root.message = msg
        root.actionText = (action && action !== "") ? action : ""
        _show()
    }

    function _show() {
        // Надежно получаем ширину монитора
        let screenW = Screen.width
        
        let startX = screenW + 50
        let targetX = screenW - root.width - 20 
        
        root.x = startX
        root.y = 60 
        root.opacity = 0.0
        
        openXAnim.from = startX
        openXAnim.to = targetX
        closeXAnim.from = targetX
        closeXAnim.to = startX
        
        root.show()
        // ВАЖНО: не вызываем requestActivate() — уведомление не должно
        // красть фокус и перебивать пользователя, пока он работает в другом окне
        openAnim.start()
    }
}