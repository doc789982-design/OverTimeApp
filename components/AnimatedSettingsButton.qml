import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

Item {
    id: root
    
    implicitWidth: 32
    implicitHeight: 32

    property color iconColor: AppTheme.textSecondary
    property color iconHoverColor: AppTheme.textPrimary
    property int bgRadius: AppTheme.radiusPill // Идеально круглая!

    signal clicked()

    Rectangle {
        anchors.fill: parent
        radius: root.bgRadius
        color: mouseArea.pressed ? AppTheme.statePress : (mouseArea.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    IconImage {
        id: gearItem
        anchors.centerIn: parent
        source: "../icons/settings.svg"
        width: AppTheme.iconLarge 
        height: AppTheme.iconLarge
        
        color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        
        transformOrigin: Item.Center
    }

    ParallelAnimation {
        id: spinAnim
        NumberAnimation { 
            target: gearItem; property: "rotation"; to: 360 
            duration: 800; easing.type: AppTheme.easeEnter
        }
        SequentialAnimation {
            NumberAnimation { target: gearItem; property: "scale"; to: 0.8; duration: AppTheme.durStandard; easing.type: AppTheme.easeExit }
            NumberAnimation { target: gearItem; property: "scale"; to: 1.0; duration: AppTheme.durSlow; easing.type: Easing.OutBack }
        }
    }

    ParallelAnimation {
        id: resetAnim
        NumberAnimation { 
            target: gearItem; property: "rotation"; to: 0
            duration: AppTheme.durSlow; easing.type: AppTheme.easeStandard 
        }
        NumberAnimation { target: gearItem; property: "scale"; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
    }

    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        
        onEntered: { resetAnim.stop(); spinAnim.restart() }
        onExited: { spinAnim.stop(); resetAnim.restart() }
        onClicked: { gearItem.rotation = 0; spinAnim.restart(); root.clicked() }
    }
}