import QtQuick

// Эффект удаления («танос»): карточка растворяется в пыли.
//
// Холст НЕ на всё окно, а по размеру карточки + запас на разлёт:
// большой Canvas при обновлении каждый кадр перезаливает текстуру
// размером с окно (см. доку Qt Quick Canvas, раздел Threaded Rendering
// and Render Target) — именно это грузило машины, а не сама пыль.
// Траектории пыли «запечены»: 10 вариантов на размер карточки,
// генерируются один раз, дальше эффект проигрывает готовый вариант.
Item {
    id: root
    anchors.fill: parent
    z: 999999

    visible: isExploding

    property bool isExploding: false
    property Item targetItem: null
    property var particles: []
    property real waveProgress: 0.0

    // Запас холста вокруг карточки: вправо пыль летит дальше всего
    readonly property real padLeft: 24
    readonly property real padTop: 80
    readonly property real padRight: 150
    readonly property real padBottom: 24

    // Запечённые траектории: { "WxH": [вариант, ...] }
    property var variantCache: ({})
    readonly property int variantLimit: 10
    readonly property int sizeLimit: 12

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

        property real ox: 0      // карточка внутри холста
        property real oy: 0
        property int iw: 0       // целый размер карточки
        property int ih: 0
        property int frame: 0

        function initExplosion() {
            var ctx = getContext("2d");
            ctx.clearRect(0, 0, width, height);

            // Целые размеры: у ячеек сетки ширина дробная, а индекс пикселя
            // в данных изображения обязан быть целым — раньше из-за этого
            // у «дробных» карточек выживала только верхняя строка пыли
            var w = canvas.iw;
            var h = canvas.ih;

            ctx.drawImage(hiddenImage, canvas.ox, canvas.oy, w, h);
            var imgData = ctx.getImageData(canvas.ox, canvas.oy, w, h);
            var data = imgData.data;

            // Готовый вариант траекторий под этот размер карточки
            var v = root.takeVariant(w, h);
            var cols = Math.ceil(w / 2);
            var rows = Math.ceil(h / 2);

            var pList = [];
            for (var row = 0; row < rows; row++) {
                var y = row * 2;
                for (var col = 0; col < cols; col++) {
                    var x = col * 2;
                    var i = row * cols + col;

                    // Пропускаем ~50% частиц — как раньше
                    if (v.skip[i]) continue;

                    var idx = (y * w + x) * 4;
                    if (data[idx + 3] > 10) {
                        pList.push({
                            currX: canvas.ox + x,
                            currY: canvas.oy + y,
                            vx: v.vx[i],
                            vy: v.vy[i],
                            z: 0.0,
                            vz: v.vz[i],
                            phaseX: v.phaseX[i],
                            phaseY: v.phaseY[i],
                            speedX: v.speedX[i],
                            speedY: v.speedY[i],
                            life: v.life[i],
                            wakeThreshold: v.wake[i],
                            colorStr: "rgb(" + data[idx] + "," + data[idx + 1] + "," + data[idx + 2] + ")",
                            active: false
                        });
                    }
                }
            }

            // По цвету: за кадр fillStyle переключается считанные разы,
            // прозрачность — через globalAlpha, без склейки строк каждый кадр
            pList.sort(function(a, b) { return a.colorStr < b.colorStr ? -1 : 1 });

            root.particles = pList;
            root.waveProgress = 0.0;
            canvas.frame = 0;

            // Защита: элемент могли уже уничтожить (обновление модели)
            if (root.targetItem) root.targetItem.opacity = 0.0;
            ctx.clearRect(0, 0, width, height);

            root.snapshotTaken();

            // ВАЖНО: возвращаем прозрачность обратно!
            // К этому моменту бэкенд уже удалил данные и обновил модель,
            // поэтому у удалённого элемента visible=false и его не видно.
            // Если этого не сделать, постоянные элементы (значок "В", статус дня)
            // навсегда останутся с opacity=0 и не появятся при повторном добавлении.
            if (root.targetItem) root.targetItem.opacity = 1.0;

            waveAnim.restart();
            renderTimer.start();
        }

        onPaint: {
            var ctx = getContext("2d");
            ctx.clearRect(0, 0, width, height);

            var baseAlpha = 1.0 - (root.waveProgress * 1.5);
            if (baseAlpha > 0) {
                ctx.globalAlpha = baseAlpha;
                ctx.drawImage(hiddenImage, canvas.ox, canvas.oy, canvas.iw, canvas.ih);
                ctx.globalAlpha = 1.0;
            }

            var pList = root.particles;
            var len = pList.length;
            var lastColor = "";

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

                    // цвет меняем только при смене, прозрачность — числом
                    if (p.colorStr !== lastColor) {
                        ctx.fillStyle = p.colorStr;
                        lastColor = p.colorStr;
                    }
                    var a = p.life > 1 ? 1 : p.life;
                    ctx.globalAlpha = a;
                    ctx.fillRect(p.currX, p.currY, 1, 1);
                }
            }
            ctx.globalAlpha = 1.0;
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

    // Вариант под размер: первые удаления наращивают пул (до 10),
    // дальше проигрываем один из готовых — случайный
    function takeVariant(w, h) {
        var key = w + "x" + h;
        var pool = variantCache[key];
        if (!pool) {
            // Кэш размеров не бесконечный: окно растянули — размеры другие
            if (Object.keys(variantCache).length >= sizeLimit) {
                delete variantCache[Object.keys(variantCache)[0]];
            }
            pool = [];
            variantCache[key] = pool;
        }
        if (pool.length < variantLimit) {
            var v = bakeVariant(w, h);
            pool.push(v);
            return v;
        }
        return pool[Math.floor(Math.random() * pool.length)];
    }

    // Один «ролик» пыли для размера карточки: сетка 2px, ~50% пропуск,
    // скорости/фазы/порог волны — те же формулы, что были в живой генерации
    function bakeVariant(w, h) {
        var cols = Math.ceil(w / 2);
        var rows = Math.ceil(h / 2);
        var n = cols * rows;
        var v = {
            skip: new Uint8Array(n),
            vx: new Float32Array(n),
            vy: new Float32Array(n),
            vz: new Float32Array(n),
            phaseX: new Float32Array(n),
            phaseY: new Float32Array(n),
            speedX: new Float32Array(n),
            speedY: new Float32Array(n),
            life: new Float32Array(n),
            wake: new Float32Array(n)
        };
        var centerX = w / 2;
        var centerY = h / 2;
        for (var row = 0; row < rows; row++) {
            var y = row * 2;
            for (var col = 0; col < cols; col++) {
                var x = col * 2;
                var i = row * cols + col;
                var dirX = (x - centerX) / w;
                var dirY = (y - centerY) / h;
                v.skip[i] = Math.random() > 0.50 ? 1 : 0;
                v.vx[i] = (dirX * (Math.random() * 3.0 + 0.5)) + (Math.random() * 1.0);
                v.vy[i] = (dirY * (Math.random() * 3.0 + 0.5)) - (Math.random() * 0.8);
                v.vz[i] = Math.random() * 0.1 + 0.02;
                v.phaseX[i] = Math.random() * Math.PI * 2;
                v.phaseY[i] = Math.random() * Math.PI * 2;
                v.speedX[i] = 0.03 + Math.random() * 0.05;
                v.speedY[i] = 0.03 + Math.random() * 0.05;
                v.life[i] = 1.0 + (Math.random() * 0.4);
                v.wake[i] = (((w - x) / w) * 0.7) + (Math.random() * 0.3);
            }
        }
        return v;
    }

    function explode(target) {
        if (!target) return;
        root.isExploding = true;
        root.targetItem = target;

        var pt = target.mapToItem(root, 0, 0);
        var ix = Math.floor(pt.x);
        var iy = Math.floor(pt.y);
        canvas.iw = Math.max(1, Math.floor(target.width));
        canvas.ih = Math.max(1, Math.floor(target.height));

        // Холст — карточка плюс запас на разлёт, в пределах окна
        var cx = Math.max(0, Math.min(root.width - 1, ix - padLeft));
        var cy = Math.max(0, Math.min(root.height - 1, iy - padTop));
        var cw = Math.max(1, Math.min(root.width - cx, canvas.iw + padLeft + padRight));
        var ch = Math.max(1, Math.min(root.height - cy, canvas.ih + padTop + padBottom));
        canvas.x = cx;
        canvas.y = cy;
        canvas.width = cw;
        canvas.height = ch;
        canvas.ox = ix - cx;
        canvas.oy = iy - cy;

        target.grabToImage(function(res) {
            hiddenImage.source = res.url;
        });
    }
}
