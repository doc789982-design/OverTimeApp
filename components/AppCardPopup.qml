import QtQuick
import QtQuick.Controls
import QtQuick.Window
import Qt5Compat.GraphicalEffects

Popup {
    id: control

    property int maxAvailableHeight: ApplicationWindow.window ? ApplicationWindow.window.height - (AppTheme.spaceL * 2) : 800
    property int desiredHeight: innerContainer.implicitHeight + (AppTheme.spaceL * 2)
    
    height: Math.min(desiredHeight, maxAvailableHeight)

    property real originY: 0
    y: ApplicationWindow.window ? Math.min(originY, ApplicationWindow.window.height - height - AppTheme.spaceM) : originY

    // ==========================================
    // МАГИЯ ИСПРАВЛЕНИЯ БАГА (Z-INDEX)
    // +10 гарантирует, что окно будет ПОВЕРХ других модалок (Настроек)
    // ==========================================
    z: AppTheme.zModal + 10

    Behavior on height { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeStandard } }
    Behavior on y { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeStandard } }

    modal: true
    dim: true // Включаем затемнение!

    // Занавес (Overlay)
    Overlay.modal: Rectangle {
        color: AppTheme.bgOverlay
        opacity: control.opened ? 1.0 : 0.0
        Behavior on opacity { 
            NumberAnimation { 
                duration: AppTheme.durStandard; 
                easing.type: control.opened ? AppTheme.easeEnter : AppTheme.easeExit 
            } 
        }
    }

    // Анимации появления
    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.95; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
        }
    }

    // ==========================================
    // ФОН И ТЕНЬ (Level 4)
    // ==========================================
    background: Rectangle {
        color: AppTheme.bgModal 
        radius: AppTheme.radiusModal
        border.color: AppTheme.borderDivider
        border.width: 1

        layer.enabled: true
        layer.effect: DropShadow {
            transparentBorder: true
            color: AppTheme.shadowColor
            radius: AppTheme.shadowL4Blur
            verticalOffset: AppTheme.shadowL4Y
            samples: 25
        }
    }

    default property alias cardContent: innerContainer.data

    contentItem: ScrollView {
        anchors.fill: parent
        clip: true
        contentWidth: availableWidth
        ScrollBar.horizontal.policy: ScrollBar.AlwaysOff 

        Column {
            id: innerContainer
            width: parent.width - (AppTheme.spaceL * 2)
            x: AppTheme.spaceL; y: AppTheme.spaceL 
            spacing: AppTheme.spaceM
        }
    }

    property real baseShakeX: 0
    function shake() {
        if (!shakeAnimation.running) {
            baseShakeX = control.x
            shakeAnimation.start()
        }
    }
    SequentialAnimation {
        id: shakeAnimation
        NumberAnimation { target: control; property: "x"; to: baseShakeX + 10; duration: 50; easing.type: Easing.OutQuad }
        NumberAnimation { target: control; property: "x"; to: baseShakeX - 10; duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: control; property: "x"; to: baseShakeX + 8;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: control; property: "x"; to: baseShakeX - 8;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: control; property: "x"; to: baseShakeX + 4;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: control; property: "x"; to: baseShakeX;      duration: 50; easing.type: Easing.OutQuad }
    }

    function showAt(callerItem, mouseX, mouseY) {
        var globalPos = callerItem.mapToItem(null, mouseX, mouseY)
        var targetX = globalPos.x
        originY = globalPos.y + AppTheme.spaceM 
        
        var windowWidth = ApplicationWindow.window.width
        if (targetX + control.width > windowWidth) targetX = windowWidth - control.width - AppTheme.spaceL
        if (targetX < 10) targetX = 10
        
        control.x = targetX
        control.open()
    }
    
    function showCentered() {
        control.x = (ApplicationWindow.window.width - control.width) / 2
        originY = (ApplicationWindow.window.height - desiredHeight) / 2 
        control.open()
    }
}