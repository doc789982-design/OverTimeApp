import QtQuick

Item {
    id: root
    width: 2

    // Мост к системе: отключается при потере фокуса или выделении текста
    property bool systemCursorVisible: parent ? parent.cursorVisible : false

    Rectangle {
        id: drawRect
        width: 2 
        height: parent.height
        transform: Translate { x: -1 }
        
        visible: root.systemCursorVisible
        
        property int cIndex: 0
        color: cIndex === 0 ? AppTheme.accentBrand : 
              (cIndex === 1 ? AppTheme.accentDanger : AppTheme.accentWarning)
              
        opacity: 1.0 
        
        // МАГИЯ: Плавная анимация изменения прозрачности (Fade-in / Fade-out)
        // Время 350мс - идеальное "дыхание" (не слишком быстро, не слишком медленно)
        Behavior on opacity { 
            NumberAnimation { duration: 350; easing.type: Easing.InOutSine } 
        }
    }

    Timer {
        id: blinkTimer
        interval: 500 // Таймер срабатывает каждые полсекунды
        running: root.systemCursorVisible
        repeat: true
        onTriggered: {
            if (drawRect.opacity > 0.5) {
                // Если горим - начинаем затухать
                drawRect.opacity = 0.0;
            } else {
                // Если потухли - меняем цвет в невидимости и начинаем разгораться
                drawRect.cIndex = (drawRect.cIndex + 1) % 3;
                drawRect.opacity = 1.0;
            }
        }
    }

    // При активной печати курсор не мигает, а горит плотным цветом
    onXChanged: resetBlink()
    onYChanged: resetBlink()
    
    onSystemCursorVisibleChanged: {
        if (systemCursorVisible) resetBlink()
        else drawRect.opacity = 0.0 // Если убрали фокус - мгновенно гасим
    }

    function resetBlink() {
        if (!systemCursorVisible) return;
        drawRect.opacity = 1.0; // Жестко зажигаем
        blinkTimer.restart();   // Сбрасываем таймер "затухания"
    }
}