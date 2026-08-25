import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

Item {
    id: root
    
    property string text: ""
    property bool isVisible: false
    property bool dropDown: false
    property int delayMs: 350   // Задержка появления: тултип не мигает при проводе мыши
    
    // МАГИЯ: Позволяет вставлять список клавиш внутрь тултипа
    default property alias content: container.data

    width: popup.width
    height: popup.height
    z: AppTheme.zTooltip

    // Показ с задержкой: наведение должно быть осознанным.
    // Ушли с кнопки — исчезло мгновенно, без таймера.
    onIsVisibleChanged: {
        if (root.isVisible) showDelayTimer.restart()
        else { showDelayTimer.stop(); popup.close() }
    }
    Timer {
        id: showDelayTimer
        interval: Math.max(0, root.delayMs)
        onTriggered: popup.open()
    }

    ToolTip {
        id: popup
        visible: false
        x: 0
        y: 0
        
        topPadding: AppTheme.spaceS
        bottomPadding: AppTheme.spaceS
        leftPadding: AppTheme.spaceM
        rightPadding: AppTheme.spaceM

        enter: Transition {
            ParallelAnimation {
                NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast }
                NumberAnimation { property: "y"; from: root.dropDown ? -4 : 4; to: 0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            }
        }
        
        background: Rectangle {
            color: AppTheme.bgElevated
            radius: AppTheme.radiusMedium
            border.color: AppTheme.borderDivider
            border.width: 1
            // Тень-картинка вместо вычисляемой (Level 3)
            AppShadow { level: 3 }
        }

        contentItem: Column {
            id: container
            spacing: AppTheme.spaceS
            
            Text {
                visible: root.text !== ""
                text: root.text
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightMedium
                horizontalAlignment: Text.AlignHCenter
            }
        }
    }
}