import QtQuick

// ============================================================
// ОБЛАКО, ПРЕВРАЩАЮЩЕЕСЯ В КОЛЬЦО — анимация загрузки обновления
//
// Воспроизводит присланный образец (гифка, 174 кадра × 30 мс):
// контур облака перетекает в кольцо и обратно, внутри облака
// живёт маленькое облачко, которое растворяется, пока форма —
// кольцо, и вырастает, когда облако «распускается» обратно.
// Цикл 5220 мс, фазы по образцу:
//   облако    750 мс  — контур с буграми, внутри облачко
//   в кольцо  300 мс  — быстрое «схлопывание» (InCubic)
//   кольцо   1750 мс  — ровное кольцо, облачка нет
//   в облако 1250 мс  — медленное «распускание» (OutCubic),
//                        облачко вырастает
//   облако   1170 мс  — покой до конца цикла
//
// Геометрия: контур облака — объединение кругов-«бугров»;
// для каждого угла от центра считается дальняя точка луча
// (радиус силуэта). Кольцо — тот же обход по углу, но радиус
// постоянный. Обе формы ОДНИМ параметром угла, поэтому морф
// — честное поточечное интерполирование, без скруток.
//
// caption — подпись под анимацией (статус загрузки).
// ============================================================
Item {
    id: root

    property int diameter: 96
    property color strokeColor: AppTheme.accentBrand
    property bool running: true
    property string caption: ""

    width: diameter
    height: diameter + (caption !== "" ? captionBlock : 0)

    readonly property real captionBlock: 26
    readonly property real loopMs: 5220

    // положение внутри цикла, мс (0..loopMs)
    property real clock: 0

    onClockChanged: canvas.requestPaint()
    onStrokeColorChanged: canvas.requestPaint()
    onDiameterChanged: canvas.requestPaint()

    NumberAnimation {
        id: loopAnim
        target: root
        property: "clock"
        from: 0
        to: root.loopMs
        duration: root.loopMs
        loops: Animation.Infinite
        running: root.running && root.visible
    }

    // ── фаза морфа: 0 = облако, 1 = кольцо ──
    function morphT(c) {
        if (c < 750)  return 0                          // облако
        if (c < 1050) {                                 // в кольцо — быстро
            var k = (c - 750) / 300
            return k * k * k
        }
        if (c < 2800) return 1                          // кольцо
        if (c < 4050) {                                 // в облако — плавно,
            var m = (c - 2800) / 1250                   // с торможением к концу
            return 1 - m * m * (3 - 2 * m)
        }
        return 0                                        // облако до цикла
    }

    // ── облачко внутри: 1 = целиком, 0 = растворилось ──
    function innerT(c) {
        if (c < 750)  return 1
        if (c < 1050) return 1 - (c - 750) / 300        // тает при схлопывании
        if (c < 2800) return 0
        if (c < 4050) return (c - 2800) / 1250          // растёт при распускании
        return 1
    }

    // радиус силуэта облака по углу: дальняя точка луча
    // с одним из бугров (объединение кругов), а снизу контур
    // подрезан прямой линией — плоское «дно» облака
    function cloudR(a) {
        var bumps = [
            [-0.335,  0.115, 0.150],   // левый маленький хвостик
            [-0.130, -0.060, 0.195],   // средняя левая туча
            [ 0.075, -0.125, 0.235],   // большая верхняя
            [ 0.290,  0.010, 0.155]    // правая
        ]
        var best = 0.20
        for (var i = 0; i < bumps.length; i++) {
            var dx = bumps[i][0], dy = bumps[i][1], r = bumps[i][2]
            var d = Math.sqrt(dx * dx + dy * dy)
            var ang = Math.atan2(dy, dx)
            var diff = a - ang
            while (diff >  Math.PI) diff -= Math.PI * 2
            while (diff < -Math.PI) diff += Math.PI * 2
            var s = d * Math.sin(diff)
            if (Math.abs(s) <= r) {
                var t = d * Math.cos(diff) + Math.sqrt(r * r - s * s)
                if (t > best) best = t
            }
        }
        // плоское дно: контур не ниже yBot (ось Y вниз)
        var sy = Math.sin(a)
        if (sy > 0.05) {
            var cap = 0.30 / sy
            if (best > cap) best = cap
        }
        return best
    }

    Canvas {
        id: canvas
        width: root.diameter
        height: root.diameter
        onPaint: {
            var ctx = getContext("2d")
            ctx.reset()
            var d = root.diameter
            var cx = d / 2, cy = d / 2
            var c = Math.min(root.clock, root.loopMs)
            var t = root.morphT(c)

            // облако чуть шире, чем выше; кольцо — круг
            var cloudSquash = 0.94
            var ringR = 0.355
            var N = 96
            var pts = []
            for (var i = 0; i <= N; i++) {
                var a = -Math.PI / 2 + (Math.PI * 2) * (i / N)
                var rc = root.cloudR(a)
                var px = cx + rc * Math.cos(a) * d * (1 - t)
                          + ringR * Math.cos(a) * d * t
                var py = cy + (rc * Math.sin(a) * cloudSquash) * d * (1 - t)
                          + ringR * Math.sin(a) * d * t
                pts.push(px, py)
            }

            ctx.lineWidth = Math.max(2, d * 0.075)
            ctx.lineCap = "round"
            ctx.lineJoin = "round"
            ctx.strokeStyle = root.strokeColor
            ctx.beginPath()
            ctx.moveTo(pts[0], pts[1])
            for (var k = 2; k < pts.length; k += 2)
                ctx.lineTo(pts[k], pts[k + 1])
            ctx.closePath()
            ctx.stroke()

            // облачко внутри — тает в кольце, дышит в облаке
            var it = root.innerT(c)
            if (it > 0.01) {
                var breath = 1 + 0.03 * Math.sin(c / 480)
                var mini = 0.36 * it * breath
                var my = cy + 0.06 * d * it     // сидит чуть ниже центра
                ctx.globalAlpha = 0.45 * it
                ctx.beginPath()
                for (var j = 0; j <= N; j++) {
                    var b = -Math.PI / 2 + (Math.PI * 2) * (j / N)
                    var rm = root.cloudR(b) * mini
                    var mx = cx + rm * Math.cos(b) * d
                    var mmy = my + rm * Math.sin(b) * cloudSquash * d
                    if (j === 0) ctx.moveTo(mx, mmy)
                    else ctx.lineTo(mx, mmy)
                }
                ctx.closePath()
                ctx.stroke()
                ctx.globalAlpha = 1
            }
        }
    }

    Text {
        id: captionText
        visible: root.caption !== ""
        y: root.diameter + 4
        width: Math.max(root.width, 160)
        x: (root.width - width) / 2
        horizontalAlignment: Text.AlignHCenter
        text: root.caption
        color: AppTheme.textTertiary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeSmall
        font.weight: AppTheme.weightMedium
        elide: Text.ElideRight
    }
}
