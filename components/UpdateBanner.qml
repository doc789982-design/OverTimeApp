import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

// Полоска «Обновить» — форма как раньше, заливка как кнопка в Telegram.
// Крестика нет: если обновление готово, полоска остаётся, пока не нажмут.
// Показываем ТОЛЬКО когда обновление готово (updateReady). Во время фоновой
// проверки/скачивания (updateBusy) полоску не показываем — чтобы ничего
// не мигало и не было видно, что программа что-то проверяет.
Item {
    id: root
    height: backend.updateReady ? 52 : 0
    visible: height > 0
    clip: true

    Behavior on height {
        NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeStandard }
    }

    Rectangle {
        id: bar
        anchors.fill: parent
        anchors.leftMargin: AppTheme.spaceM
        anchors.rightMargin: AppTheme.spaceM
        anchors.bottomMargin: AppTheme.spaceXS
        radius: AppTheme.radiusMedium

        gradient: Gradient {
            orientation: Gradient.Horizontal
            GradientStop { position: 0.0; color: "#2BD16A" }
            GradientStop { position: 1.0; color: "#18C8C8" }
        }

        // Тёмная подложка поверх градиента — один и тот же «приглушённый» вид
        // и когда обновление готовится, и когда ждёт нажатия «Обновить».
        Rectangle {
            anchors.fill: parent
            radius: parent.radius
            color: Qt.rgba(0, 0, 0, 0.18)
        }

        Rectangle {
            anchors.fill: parent
            radius: parent.radius
            visible: pressArea.containsMouse && !backend.updateBusy
            color: Qt.rgba(1, 1, 1, pressArea.pressed ? 0.14 : 0.08)
        }

        Row {
            anchors.verticalCenter: parent.verticalCenter
            anchors.left: parent.left
            anchors.leftMargin: AppTheme.spaceM
            anchors.right: parent.right
            anchors.rightMargin: AppTheme.spaceM
            spacing: AppTheme.spaceS

            Rectangle {
                width: 28
                height: 28
                radius: 14
                anchors.verticalCenter: parent.verticalCenter
                color: Qt.rgba(1, 1, 1, 0.22)

                IconImage {
                    anchors.centerIn: parent
                    source: "../icons/refresh.svg"
                    width: AppTheme.iconMedium
                    height: AppTheme.iconMedium
                    color: "#FFFFFF"
                }
            }

            Text {
                anchors.verticalCenter: parent.verticalCenter
                width: parent.width - 28 - parent.spacing
                elide: Text.ElideRight
                text: backend.updateBusy
                      ? (backend.updateStatusText || "Готовим обновление…")
                      : "Обновить OverTimeTab"
                color: "#FFFFFF"
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                font.weight: AppTheme.weightBold
            }
        }

        MouseArea {
            id: pressArea
            anchors.fill: parent
            enabled: backend.updateReady && !backend.updateBusy
            hoverEnabled: true
            cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
            onClicked: backend.applyReadyUpdate()
        }
    }
}
