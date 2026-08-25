import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

// Полоска «Обновить» — как в Telegram: снизу слева, не перекрывает работу.
Item {
    id: root
    height: (backend.updateReady || backend.updateBusy) ? 52 : 0
    visible: height > 0
    clip: true

    Behavior on height {
        NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeStandard }
    }

    Rectangle {
        anchors.fill: parent
        anchors.leftMargin: AppTheme.spaceM
        anchors.rightMargin: AppTheme.spaceM
        anchors.bottomMargin: AppTheme.spaceXS
        radius: AppTheme.radiusMedium
        color: backend.updateBusy ? AppTheme.bgElevated : AppTheme.accentBrand
        border.color: backend.updateBusy ? AppTheme.borderDivider : "transparent"
        border.width: 1

        Behavior on color { ColorAnimation { duration: AppTheme.durFast } }

        Row {
            anchors.fill: parent
            anchors.leftMargin: AppTheme.spaceM
            anchors.rightMargin: AppTheme.spaceS
            spacing: AppTheme.spaceS

            Item {
                width: parent.width - (dismissBtn.visible ? dismissBtn.width + parent.spacing : 0)
                height: parent.height

                Text {
                    anchors.verticalCenter: parent.verticalCenter
                    anchors.left: parent.left
                    anchors.right: parent.right
                    elide: Text.ElideRight
                    text: backend.updateBusy
                          ? (backend.updateStatusText || "Готовим обновление…")
                          : (backend.updateVersion
                             ? ("Обновить до " + backend.updateVersion)
                             : "Обновить программу")
                    color: backend.updateBusy ? AppTheme.textPrimary : AppTheme.textOnAccent
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                }

                MouseArea {
                    anchors.fill: parent
                    enabled: backend.updateReady && !backend.updateBusy
                    cursorShape: enabled ? Qt.PointingHandCursor : Qt.ArrowCursor
                    onClicked: backend.applyReadyUpdate()
                    hoverEnabled: true
                }
            }

            Rectangle {
                id: dismissBtn
                visible: backend.updateReady && !backend.updateBusy
                width: 28
                height: 28
                radius: 14
                anchors.verticalCenter: parent.verticalCenter
                color: dismissHov.containsMouse ? Qt.rgba(1, 1, 1, 0.18) : "transparent"

                Text {
                    anchors.centerIn: parent
                    text: "✕"
                    color: AppTheme.textOnAccent
                    font.pixelSize: 12
                    font.weight: AppTheme.weightBold
                }
                MouseArea {
                    id: dismissHov
                    anchors.fill: parent
                    hoverEnabled: true
                    cursorShape: Qt.PointingHandCursor
                    onClicked: backend.dismissUpdate()
                }
            }
        }
    }
}
