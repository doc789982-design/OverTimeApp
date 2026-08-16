import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

Item {
    id: root
    
    implicitWidth: 32
    implicitHeight: 32

    property string iconSource: ""
    property int iconSize: AppTheme.iconMedium
    property color iconColor: AppTheme.textSecondary
    property color iconHoverColor: AppTheme.textPrimary
    property int bgRadius: AppTheme.radiusMedium

    signal clicked()

    scale: mouseArea.pressed ? AppTheme.scaleActive : 1.0
    Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }

    Rectangle {
        id: bg
        anchors.fill: parent
        radius: root.bgRadius
        color: mouseArea.pressed ? AppTheme.statePress : (mouseArea.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    IconImage {
        id: iconItem
        anchors.centerIn: parent
        source: root.iconSource
        width: root.iconSize
        height: root.iconSize
        color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        onClicked: root.clicked()
    }
}