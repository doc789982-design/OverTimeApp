import QtQuick

// Эффект удаления («танос»): карточка растворяется в пыли.
//
// Два пути:
//  1. GPU (ShaderEffect) — основная дорога, идея из Telegram Desktop
//     (ui/effects/thanos_effect): случайность не хранят, а вычисляют
//     хешем от координаты и зерна, цвет пылинка берёт из самого снимка
//     карточки. CPU в кадре не считает ничего — только время.
//  2. Canvas — запасной путь для машин без графического ускорителя
//     (и для песочницы тестов): холст по размеру карточки, траектории
//     «запечены» (10 вариантов на размер карточки).
//
// Общие для обоих путей: запас холста вокруг карточки (вправо пыль летит
// дальше всего), волна справа налево, время жизни пылинок 1.39–1.94 с.
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

    // Запечённые траектории запасного пути: { "WxH": [вариант, ...] }
    property var variantCache: ({})
    readonly property int variantLimit: 10
    readonly property int sizeLimit: 12

    // ── GPU-путь ──
    // Графическое ускорение есть? Software/Null/Unknown — нет,
    // идём запасным путём (имена значений — из доки GraphicsInfo)
    readonly property bool gpuOK: GraphicsInfo.api !== GraphicsInfo.Software
                                  && GraphicsInfo.api !== GraphicsInfo.Null
                                  && GraphicsInfo.api !== GraphicsInfo.Unknown
    property real tSec: 0              // время эффекта, с
    property real seed: 0              // зерно пыли: каждый взрыв новый
    property bool snapshotReady: false
    property bool baseReady: false
    property bool started: false
    // волна 0.5 с + максимальная жизнь пылинки (1.389 + 0.555 с)
    readonly property real waveSec: 0.5
    readonly property real totalSec: waveSec + 1.389 + 0.555 + 0.05
    // доля волны — кривая торможения OutQuad, как в запасном пути
    readonly property real waveFrac: {
        var p = tSec / waveSec
        if (p > 1) p = 1
        return 1 - (1 - p) * (1 - p)
    }
    // классы пыли: скорость семьи, px/с (vx, vy) + постоянный снос (вправо-вверх)
    readonly property var dustMotions: [
        Qt.vector4d(55, -30, 72, -24),
        Qt.vector4d(90, -55, 72, -24),
        Qt.vector4d(35, -10, 72, -24)
    ]
    // размер пылинок по семьям, px (быстрая — мелкая, ленивая — крупная)
    // и сколько пылинок на ячейку сетки: мелкой пыли больше, крупной — меньше
    readonly property var dustSizes: [3.0, 2.3, 4.0]
    readonly property var dustDensities: [1.7, 2.0, 1.2]
    readonly property real dustSizeJitter: 1.1
    readonly property real dustGravity: 10

    signal finished()
    signal snapshotTaken()

    Image {
        id: hiddenImage
        visible: false
        onStatusChanged: {
            if (status === Image.Ready || status === Image.Error) {
                root.snapshotReady = true
                root.tryBegin()
            }
        }
    }

    // ── GPU-путь: гаснущий остаток карточки + три слоя пыли ──
    Item {
        id: shaderStage
        visible: root.isExploding && root.gpuOK
        clip: true

        Image {
            id: baseImage
            // остаток карточки гаснет так же, как в запасном пути:
            // общая прозрачность 1 − волна×1.5
            opacity: Math.max(0, 1 - root.waveFrac * 1.5)
            fillMode: Image.Stretch
            onStatusChanged: {
                if (status === Image.Ready || status === Image.Error) {
                    root.baseReady = true
                    root.tryBegin()
                }
            }
        }

        // Пыль создаём только при графическом ускорении: под софтверным
        // рендером ShaderEffect не работает и не должен даже создаваться
        Loader {
            id: dustLoader
            anchors.fill: parent
            active: root.gpuOK
            sourceComponent: Component {
                Repeater {
                    model: root.dustMotions.length
                    ShaderEffect {
                        anchors.fill: parent
                        blending: true
                        property var src: hiddenImage
                        property real uT: root.tSec
                        // зерно у каждого слоя своё: иначе семьи дадут
                        // коррелированную решётку одинаковых пылинок
                        property real uSeed: root.seed + index * 0.7371
                        property vector2d uQuad: Qt.vector2d(shaderStage.width, shaderStage.height)
                        property vector4d uCard: Qt.vector4d(baseImage.x, baseImage.y,
                                                             baseImage.width, baseImage.height)
                        property vector4d uMotion: root.dustMotions[index]
                        property vector4d uClass: Qt.vector4d(root.dustDensities[index],
                                                              root.dustGravity,
                                                              root.dustSizes[index],
                                                              root.dustSizeJitter)
                        property vector2d uLife: Qt.vector2d(1.389, 0.555)
                        fragmentShader: Qt.resolvedUrl("../shaders/thanos_dust.frag.qsb")
                    }
                }
            }
        }
    }

    Canvas {
        id: canvas

        property real itemX: 0
        property real itemY: 0
        property int iw: 0
        property int ih: 0
        property int frame: 0

        visible: !root.gpuOK

        function initExplosion() {
            var ctx = getContext("2d");
            ctx.clearRect(0, 0, width, height);

            // Целые размеры: у ячеек сетки ширина дробная, а индекс пикселя
            // в данных изображения обязан быть целым — раньше из-за этого
            // у «дробных» карточек выживала только верхняя строка пыли
            var w = canvas.iw;
            var h = canvas.ih;

            ctx.drawImage(hiddenImage, canvas.itemX, canvas.itemY, w, h);
            var imgData = ctx.getImageData(canvas.itemX, canvas.itemY, w, h);
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
                            currX: canvas.itemX + x,
                            currY: canvas.itemY + y,
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
                ctx.drawImage(hiddenImage, canvas.itemX, canvas.itemY, canvas.iw, canvas.ih);
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
                    // пылинка крупная и к концу жизни сжимается — как в шейдере
                    var s = 0.8 + 1.8 * a;
                    ctx.fillRect(p.currX, p.currY, s, s);
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

    // ── GPU-путь: часы эффекта (линейное время, волна изгибается
    // в шейдере и в waveFrac) ──
    NumberAnimation {
        id: tAnim
        target: root
        property: "tSec"
        from: 0
        to: root.totalSec
        duration: root.totalSec * 1000
        onStopped: {
            root.isExploding = false
            root.finished()
        }
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

    // Старт эффекта: оба изображения готовы (снимок и его копия для базы)
    function tryBegin() {
        if (!isExploding || started || !snapshotReady || !baseReady) return
        started = true
        if (gpuOK) {
            if (root.targetItem) root.targetItem.opacity = 0.0
            root.snapshotTaken()
            if (root.targetItem) root.targetItem.opacity = 1.0
            root.seed = Math.random()
            root.tSec = 0
            tAnim.restart()
        } else {
            canvas.initExplosion()
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
        root.started = false;
        root.snapshotReady = false;
        root.baseReady = false;

        var pt = target.mapToItem(root, 0, 0);
        var ix = Math.floor(pt.x);
        var iy = Math.floor(pt.y);
        canvas.iw = Math.max(1, Math.floor(target.width));
        canvas.ih = Math.max(1, Math.floor(target.height));

        // Холст — карточка плюс запас на разлёт, в пределах окна.
        // Один и тот же прямоугольник у обоих путей: Canvas и шейдерный слой
        var cx = Math.max(0, Math.min(root.width - 1, ix - padLeft));
        var cy = Math.max(0, Math.min(root.height - 1, iy - padTop));
        var cw = Math.max(1, Math.min(root.width - cx, canvas.iw + padLeft + padRight));
        var ch = Math.max(1, Math.min(root.height - cy, canvas.ih + padTop + padBottom));

        canvas.x = cx;
        canvas.y = cy;
        canvas.width = cw;
        canvas.height = ch;
        canvas.itemX = ix - cx;
        canvas.itemY = iy - cy;

        shaderStage.x = cx;
        shaderStage.y = cy;
        shaderStage.width = cw;
        shaderStage.height = ch;
        baseImage.x = ix - cx;
        baseImage.y = iy - cy;
        baseImage.width = canvas.iw;
        baseImage.height = canvas.ih;

        // База GPU-пути грузит тот же снимок; пыль сэмплерирует hiddenImage
        baseImage.source = "";

        target.grabToImage(function(res) {
            hiddenImage.source = res.url;
            baseImage.source = res.url;
        });
    }
}
