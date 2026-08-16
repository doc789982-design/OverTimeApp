import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

Item {
    id: root
    
    implicitWidth: 32
    implicitHeight: 32

    property color iconColor: AppTheme.textSecondary
    property color iconHoverColor: AppTheme.textPrimary
    property int bgRadius: AppTheme.radiusMedium

    signal clicked()

    Rectangle {
        anchors.fill: parent
        radius: root.bgRadius
        color: mouseArea.pressed ? AppTheme.statePress : (mouseArea.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    Item {
        width: AppTheme.iconLarge
        height: AppTheme.iconLarge
        anchors.centerIn: parent

        Item {
            id: clipBox
            x: 0; y: -20; width: 20; height: 27; clip: true 

            IconImage {
                id: paperLayer
                source: "../icons/print_paper.svg"
                width: AppTheme.iconLarge; height: AppTheme.iconLarge
                y: 19 
                color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
            }
        }

        IconImage {
            id: bodyLayer
            source: "../icons/print_body.svg"
            width: AppTheme.iconLarge; height: AppTheme.iconLarge
            color: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
            transformOrigin: Item.Bottom
        }
    }

    SequentialAnimation {
        id: printAnim
        
        ParallelAnimation {
            NumberAnimation { target: paperLayer; property: "y"; to: -2; duration: AppTheme.durStandard; easing.type: AppTheme.easeExit }
            NumberAnimation { target: paperLayer; property: "opacity"; to: 0.0; duration: AppTheme.durStandard }
            
            SequentialAnimation {
                NumberAnimation { target: bodyLayer; property: "scale"; to: 0.85; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
                NumberAnimation { target: bodyLayer; property: "scale"; to: 1.0; duration: AppTheme.durFast; easing.type: Easing.OutBounce }
            }
        }
        
        PropertyAction { target: paperLayer; property: "y"; value: 27 }
        
        ParallelAnimation {
            NumberAnimation { target: paperLayer; property: "opacity"; to: 1.0; duration: AppTheme.durMicro }
            NumberAnimation { target: paperLayer; property: "y"; to: 19; duration: AppTheme.durSlow; easing.type: Easing.OutBack }
        }
    }

    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        
        onEntered: { if (!printAnim.running) printAnim.start() }
        onClicked: { if (!printAnim.running) printAnim.start(); root.clicked() }
    }
}