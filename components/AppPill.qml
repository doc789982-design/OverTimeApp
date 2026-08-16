import QtQuick
import QtQuick.Controls

Item {
    id: root
    
    implicitHeight: 28
    implicitWidth: layoutRow.implicitWidth + 24

    property string text: ""
    property bool isAction: false    
    property bool removable: false   

    signal clicked()
    signal removeClicked()

    Rectangle {
        anchors.fill: parent
        radius: AppTheme.radiusSmall // 4px (или 6px, если хочешь чуть мягче)
        
        // МАГИЯ СТРОГИХ ЦВЕТОВ:
        // Используем новые стейты (statePress / stateHover) и новые границы
        color: mainMouse.pressed ? AppTheme.statePress : 
               (mainMouse.containsMouse ? AppTheme.stateHover : 
               (root.isAction ? "transparent" : AppTheme.bgInput))
               
        border.color: root.isAction && mainMouse.containsMouse ? AppTheme.accentBrand : AppTheme.borderDivider
        border.width: 1
        
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }

        Row {
            id: layoutRow
            anchors.centerIn: parent
            spacing: AppTheme.spaceXS // Системный отступ (8px)

            Text {
                visible: root.isAction
                text: "+"
                color: root.isAction ? AppTheme.accentBrand : AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBodyLarge
                font.weight: AppTheme.weightBold
                anchors.verticalCenter: parent.verticalCenter
                anchors.verticalCenterOffset: -1 
            }

            Text {
                text: root.text
                color: root.isAction ? AppTheme.accentBrand : AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                font.weight: root.isAction ? AppTheme.weightBold : AppTheme.weightRegular
                anchors.verticalCenter: parent.verticalCenter
            }

            // Крестик удаления
            Item {
                visible: root.removable
                width: 16
                height: 16
                anchors.verticalCenter: parent.verticalCenter

                Rectangle {
                    anchors.centerIn: parent
                    width: 16; height: 16; radius: 8
                    
                    // Цвет крестика: прижато -> ярко-красный, наведено -> мягкий красный фон
                    color: removeMouse.pressed ? AppTheme.accentDanger : 
                           (removeMouse.containsMouse ? AppTheme.bgDangerSoft : AppTheme.stateHover)
                    Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                    Text {
                        anchors.centerIn: parent
                        text: "✕"
                        // Если навели на кружок - крестик становится красным, иначе он просто серый
                        color: removeMouse.containsMouse ? AppTheme.accentDanger : AppTheme.textSecondary
                        font.pixelSize: 9
                        font.weight: AppTheme.weightBold
                    }
                }

                MouseArea {
                    id: removeMouse
                    anchors.fill: parent
                    anchors.margins: -4 
                    hoverEnabled: true
                    onClicked: root.removeClicked()
                }
            }
        }

        MouseArea {
            id: mainMouse
            anchors.left: parent.left
            anchors.top: parent.top
            anchors.bottom: parent.bottom
            anchors.right: root.removable ? layoutRow.right : parent.right
            anchors.rightMargin: root.removable ? 24 : 0
            hoverEnabled: true
            onClicked: root.clicked()
        }
    }
}