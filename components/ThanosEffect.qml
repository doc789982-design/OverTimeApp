import QtQuick

Item {
    id: root
    anchors.fill: parent
    z: 999999

    visible: isExploding

    property bool isExploding: false
    property Item targetItem: null
    property var particles: []
    property real waveProgress: 0.0

    signal finished()
    signal snapshotTaken()

    Image {
        id: hiddenImage
        visible: false
        onStatusChanged: {
            if (status === Image.Ready) {
                canvas.initExplosion();
            }
        }
    }

    Canvas {
        id: canvas
        anchors.fill: parent
        
        property real itemX: 0
        property real itemY: 0
        property real itemW: 0
        property real itemH: 0
        property int frame: 0 

        function initExplosion() {
            var ctx = getContext("2d");
            ctx.clearRect(0, 0, width, height);
            
            ctx.drawImage(hiddenImage, itemX, itemY, itemW, itemH);
            var imgData = ctx.getImageData(itemX, itemY, itemW, itemH);
            var data = imgData.data;
            
            var pList = [];
            var centerX = itemW / 2;
            var centerY = itemH / 2;
            
            // Шаг 2px — в 4 раза меньше частиц, визуально незаметно
            var step = 2;
            
            for (var y = 0; y < itemH; y += step) {
                for (var x = 0; x < itemW; x += step) {
                    
                    // Пропускаем ~50% частиц
                    if (Math.random() > 0.50) continue; 
                    
                    var idx = (y * itemW + x) * 4;
                    if (data[idx+3] > 10) { 
                        
                        var dirX = (x - centerX) / itemW; 
                        var dirY = (y - centerY) / itemH;
                        
                        var normalizedX = (itemW - x) / itemW; 
                        var wakeThreshold = (normalizedX * 0.7) + (Math.random() * 0.3);
                        
                        var colorPrefix = "rgba(" + data[idx] + "," + data[idx+1] + "," + data[idx+2] + ",";
                        
                        pList.push({
                            currX: itemX + x,
                            currY: itemY + y,
                            vx: (dirX * (Math.random() * 3.0 + 0.5)) + (Math.random() * 1.0),
                            vy: (dirY * (Math.random() * 3.0 + 0.5)) - (Math.random() * 0.8),
                            z: 0.0,
                            vz: Math.random() * 0.1 + 0.02,
                            phaseX: Math.random() * Math.PI * 2,
                            phaseY: Math.random() * Math.PI * 2,
                            speedX: 0.03 + Math.random() * 0.05,
                            speedY: 0.03 + Math.random() * 0.05,
                            life: 1.0 + (Math.random() * 0.4), 
                            wakeThreshold: wakeThreshold,
                            colorStr: colorPrefix,
                            active: false
                        });
                    }
                }
            }
            
            root.particles = pList;
            root.waveProgress = 0.0;
            canvas.frame = 0;
            
            root.targetItem.opacity = 0.0;
            ctx.clearRect(0, 0, width, height); 

            root.snapshotTaken();
            
            waveAnim.restart();
            renderTimer.start();
        }

        onPaint: {
            var ctx = getContext("2d");
            ctx.clearRect(0, 0, width, height);
            
            var baseAlpha = 1.0 - (root.waveProgress * 1.5);
            if (baseAlpha > 0) {
                ctx.globalAlpha = baseAlpha;
                ctx.drawImage(hiddenImage, itemX, itemY, itemW, itemH);
                ctx.globalAlpha = 1.0;
            }

            var pList = root.particles;
            var len = pList.length;

            for (var i = 0; i < len; i++) {
                var p = pList[i];
                
                if (!p.active && root.waveProgress >= p.wakeThreshold) {
                    p.active = true;
                }
                
                if (p.active && p.life > 0) {
                    
                    p.z += p.vz;
                    p.vz -= 0.01;
                    if (p.z < 0) p.z = 0;
                    
                    var perspective = 1.0 + p.z; 
                    
                    p.currX += p.vx * perspective;
                    p.currY += p.vy * perspective;
                    
                    p.vx *= 0.88; 
                    p.vy *= 0.88;
                    
                    p.currX += 1.2;
                    p.currY -= 0.4; 
                    
                    p.currX += Math.sin(canvas.frame * p.speedX + p.phaseX) * 0.5;
                    p.currY += Math.cos(canvas.frame * p.speedY + p.phaseY) * 0.5;
                    
                    p.life -= 0.012; 
                    
                    ctx.fillStyle = p.colorStr + p.life + ")";
                    ctx.fillRect(p.currX, p.currY, 1, 1); 
                }
            }
        }
    }

    NumberAnimation on waveProgress {
        id: waveAnim
        from: 0.0
        to: 1.0
        duration: 500 
        easing.type: Easing.OutQuad
    }

    Timer {
        id: renderTimer
        interval: 16 
        repeat: true
        onTriggered: {
            canvas.frame += 1;
            canvas.requestPaint();
            
            if (waveAnim.running === false) {
                var allDead = true;
                var pList = root.particles;
                for (var i = 0; i < pList.length; i++) {
                    if (pList[i].life > 0) {
                        allDead = false;
                        break;
                    }
                }
                if (allDead) {
                    renderTimer.stop();
                    root.particles = [];
                    root.isExploding = false;
                    root.finished();
                }
            }
        }
    }

    function explode(target) {
        if (!target) return;
        root.isExploding = true;
        root.targetItem = target;
        
        var pt = target.mapToItem(root, 0, 0);
        canvas.itemX = pt.x;
        canvas.itemY = pt.y;
        canvas.itemW = target.width;
        canvas.itemH = target.height;
        
        target.grabToImage(function(res) {
            hiddenImage.source = res.url;
        });
    }
}
