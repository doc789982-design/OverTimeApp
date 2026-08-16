import QtQuick

Item {
    id: root
    anchors.fill: parent
    z: 999998 
    
    // Эффект видим, если идет анимация ИЛИ если мы "подготавливаем" перекрытие
    visible: isAnimating || isPreparing

    property bool isAnimating: false
    property bool isPreparing: false 
    property real clickX: 0
    property real clickY: 0
    property real currentRadius: 0
    property real maxRadius: 0
    
    property var switchThemeCallback: null
    property Item targetItem: null

    // Невидимая картинка для хранения старого экрана
    Image {
        id: screenshot
        visible: false
        
        // Как только картинка загрузилась в память...
        onStatusChanged: {
            if (status === Image.Ready && root.isPreparing) {
                canvas.requestPaint(); // 1. Принудительно рисуем сплошной холст
                delayTimer.start();    // 2. Ждем 1 кадр, чтобы видеокарта его вывела
            }
        }
    }

    // Тот самый таймер "затвора"
    Timer {
        id: delayTimer
        interval: 30 // Ждем 30мс (примерно 2 кадра при 60fps)
        repeat: false
        onTriggered: {
            // Вот теперь экран гарантированно перекрыт старой картинкой!
            root.isPreparing = false;
            root.isAnimating = true;
            
            // 3. Безопасно переключаем настоящую тему под картинкой
            root.switchThemeCallback(); 
            
            // 4. Запускаем красивое раскрытие
            radiusAnim.start();
        }
    }

    // Холст
    Canvas {
        id: canvas
        anchors.fill: parent
        onPaint: {
            var ctx = getContext("2d");
            ctx.clearRect(0, 0, width, height);
            
            // Если мы только перекрываем экран (isPreparing), рисуем СПЛОШНУЮ картинку без дыры
            if (root.isPreparing) {
                ctx.drawImage(screenshot, 0, 0, width, height);
                return;
            }
            
            // Если идет анимация, рисуем картинку и прорезаем в ней дыру
            ctx.drawImage(screenshot, 0, 0, width, height);
            ctx.globalCompositeOperation = "destination-out";
            ctx.beginPath();
            ctx.arc(root.clickX, root.clickY, root.currentRadius, 0, Math.PI * 2);
            ctx.fill();
            ctx.globalCompositeOperation = "source-over";
        }
    }

    // Анимация роста радиуса круга
    NumberAnimation {
        id: radiusAnim
        target: root
        property: "currentRadius"
        from: 0
        to: root.maxRadius
        duration: 500 
        easing.type: Easing.InOutCubic 
        
        onStopped: {
            root.isAnimating = false;
        }
    }

    onCurrentRadiusChanged: {
        if (isAnimating) canvas.requestPaint();
    }

    // Главная функция запуска
    function execute(startX, startY, callback) {
        if (isAnimating || isPreparing) return;
        
        root.clickX = startX;
        root.clickY = startY;
        root.switchThemeCallback = callback;
        
        let dx = Math.max(startX, width - startX);
        let dy = Math.max(startY, height - startY);
        root.maxRadius = Math.sqrt(dx * dx + dy * dy) + 50;
        
        root.currentRadius = 0;
        
        // Делаем мгновенный скриншот нашей обертки
        targetItem.grabToImage(function(result) {
            root.isPreparing = true;         // Делаем холст видимым
            screenshot.source = result.url;  // Передаем картинку, что запустит onStatusChanged
        });
    }
}