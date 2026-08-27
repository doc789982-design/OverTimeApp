import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Window
import Qt5Compat.GraphicalEffects

// Плашка у края экрана. Текст на всю ширину, кнопки снизу —
// иначе RowLayout сжимает абзац в столбик по одному слову.
Window {
    id: root

    flags: Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
    color: "transparent"

    width: 400
    height: card.height + 32

    property string message: "Не забудьте сделать табель, ознакомить с ним всех сотрудников и сдать его в кадровое подразделение до 5-го числа месяца. Скрыть до следующего напоминания."
    property string actionText: ""

    signal actionTriggered()

    Rectangle {
        id: card
        x: 16
        y: 16
        width: parent.width - 32
        height: inner.implicitHeight + AppTheme.spaceL * 2
        color: AppTheme.bgElevated
        radius: AppTheme.radiusMedium
        border.color: AppTheme.borderDivider
        border.width: 1

        AppShadow { level: 5; yOffset: 4 }

        Column {
            id: inner
            x: AppTheme.spaceM
            y: AppTheme.spaceM
            width: parent.width - AppTheme.spaceM * 2
            spacing: AppTheme.spaceS

            Text {
                width: parent.width
                text: root.message
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                wrapMode: Text.WordWrap
                lineHeight: 1.35
            }

            Item {
                width: parent.width
                height: btnRow.implicitHeight

                Row {
                    id: btnRow
                    anchors.right: parent.right
                    spacing: AppTheme.spaceXS

                    AppButton {
                        text: "Ок"
                        variant: "secondary"
                        implicitHeight: 32
                        onClicked: closeAnim.start()
                    }
                    AppButton {
                        visible: root.actionText !== ""
                        text: root.actionText
                        variant: "primary"
                        implicitHeight: 32
                        onClicked: { root.actionTriggered(); closeAnim.start() }
                    }
                }
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
        root.actionText = ""
        _show()
    }

    function showCustom(msg, action) {
        root.message = msg
        root.actionText = (action && action !== "") ? action : ""
        _show()
    }

    function _show() {
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
        openAnim.start()
    }
}
