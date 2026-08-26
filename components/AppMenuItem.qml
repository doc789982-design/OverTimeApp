import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

MenuItem {
    id: control

    property bool isDanger: false
    property string customColor: ""
    property string iconSource: ""
    property bool showDelete: false

    signal deleteClicked()

    padding: 0
    leftPadding: 0
    rightPadding: 0
    topPadding: 0
    bottomPadding: 0
    implicitHeight: visible ? 36 : 0
    implicitWidth: visible ? contentItem.implicitWidth : 0
    opacity: enabled ? 1.0 : AppTheme.alphaDisabled
    hoverEnabled: true

    indicator: Item { implicitWidth: 0; implicitHeight: 0 }
    arrow: Item { implicitWidth: 0; implicitHeight: 0 }

    readonly property color _ink: customColor !== "" ? customColor
                                : (isDanger ? AppTheme.accentDanger : AppTheme.textPrimary)
    readonly property color _iconInk: customColor !== "" ? customColor
                                    : (isDanger ? AppTheme.accentDanger : AppTheme.textSecondary)

    contentItem: Item {
        implicitHeight: 36
        implicitWidth: AppTheme.spaceS + AppTheme.iconMedium + AppTheme.spaceS
                       + Math.ceil(itemLabel.implicitWidth)
                       + (control.showDelete ? 36 : AppTheme.spaceS)

        Item {
            id: iconSlot
            width: AppTheme.iconMedium
            height: AppTheme.iconMedium
            x: AppTheme.spaceS
            anchors.verticalCenter: parent.verticalCenter

            IconImage {
                anchors.centerIn: parent
                visible: control.iconSource !== ""
                source: control.iconSource
                width: AppTheme.iconMedium
                height: AppTheme.iconMedium
                color: control._iconInk
            }
        }

        Text {
            id: itemLabel
            text: control.text
            color: control._ink
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody
            font.weight: control.isDanger ? AppTheme.weightMedium : AppTheme.weightRegular
            elide: Text.ElideRight
            verticalAlignment: Text.AlignVCenter
            anchors.verticalCenter: parent.verticalCenter
            anchors.left: parent.left
            anchors.leftMargin: AppTheme.spaceS + AppTheme.iconMedium + AppTheme.spaceS
            anchors.right: parent.right
            anchors.rightMargin: control.showDelete ? 40 : AppTheme.spaceS
        }

        Rectangle {
            id: deleteBtn
            visible: control.showDelete
            width: 28
            height: 28
            radius: AppTheme.radiusSmall
            anchors.verticalCenter: parent.verticalCenter
            anchors.right: parent.right
            anchors.rightMargin: 6
            color: deleteMouse.containsMouse ? AppTheme.bgDangerSoft : "transparent"

            IconImage {
                anchors.centerIn: parent
                source: "../icons/trash.svg"
                width: AppTheme.iconMedium
                height: AppTheme.iconMedium
                color: deleteMouse.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary
            }

            MouseArea {
                id: deleteMouse
                anchors.fill: parent
                hoverEnabled: true
                preventStealing: true
                cursorShape: Qt.PointingHandCursor
                onPressed: (mouse) => { mouse.accepted = true }
                onClicked: (mouse) => {
                    control.deleteClicked()
                    mouse.accepted = true
                }
            }
        }
    }

    background: Rectangle {
        implicitWidth: 0
        implicitHeight: 36
        color: control.hovered && !(control.showDelete && deleteMouse.containsMouse)
               ? AppTheme.stateHover : "transparent"
        radius: AppTheme.radiusSmall
        anchors.leftMargin: AppTheme.spaceXXS
        anchors.rightMargin: AppTheme.spaceXXS
        anchors.topMargin: 1
        anchors.bottomMargin: 1
    }
}
