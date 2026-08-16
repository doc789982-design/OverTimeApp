import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

Popup {
    id: root

    property string title: "Инспектор"
    default property alias sheetContent: contentArea.data

    modal: false 
    focus: true
    z: AppTheme.zModal

    // АНИМАЦИИ ВЫЕЗДА СНИЗУ
    enter: Transition { 
        NumberAnimation { property: "y"; from: parent.height; to: parent.height - root.height; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter } 
    }
    exit: Transition { 
        NumberAnimation { property: "y"; from: parent.height - root.height; to: parent.height; duration: AppTheme.durFast; easing.type: AppTheme.easeExit } 
    }

    // ФОН И ТЕНЬ (Level 4)
    background: Rectangle {
        color: AppTheme.bgModal 
        border.color: AppTheme.borderDivider
        border.width: 1
        
        // Скругляем только ВЕРХНИЕ углы
        radius: AppTheme.radiusModal
        Rectangle { 
            anchors.bottom: parent.bottom; width: parent.width; height: AppTheme.radiusModal; color: AppTheme.bgModal 
        } 
        
        layer.enabled: true
        layer.effect: DropShadow {
            transparentBorder: true
            color: AppTheme.shadowColor
            radius: AppTheme.shadowL4Blur
            verticalOffset: -AppTheme.shadowL4Y // Тень падает ВВЕРХ
            samples: 25
        }
    }

    contentItem: ColumnLayout {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceL
        spacing: AppTheme.spaceL

        // ШАПКА
        RowLayout {
            Layout.fillWidth: true
            Text {
                Layout.fillWidth: true; text: root.title
                color: AppTheme.textPrimary; font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH2; font.weight: AppTheme.weightBold
            }
            Rectangle {
                Layout.preferredWidth: 32; Layout.preferredHeight: 32; radius: AppTheme.radiusPill
                color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                IconImage { anchors.centerIn: parent; source: "../icons/close.svg"; width: AppTheme.iconSmall; height: AppTheme.iconSmall; color: AppTheme.textSecondary }
                MouseArea { id: closeHov; anchors.fill: parent; hoverEnabled: true; onClicked: root.close() }
            }
        }

        Item {
            Layout.fillWidth: true; Layout.fillHeight: true
            Item { id: contentArea; anchors.fill: parent }
        }
    }
}