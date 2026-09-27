import QtQuick

// ============================================================
// МОРФИРУЮЩАЯ ИКОНКА КНОПКИ ОБНОВЛЕНИЯ
//
// Одна геометрия ПЛАВНО перетекает в другую — без подмены
// картинок: формы («refresh», «download», «ring», «install»)
// описаны одинаковым набором точек (три штриха: стержень,
// наконечник, поднос), и интерполируются параметром morph.
//
//   refresh  — дуга с стрелкой («перезагрузка/проверка»)
//   download — стрелка вниз на поднос («скачать»)
//   ring     — кольцо, заполняющееся progress («загрузка»)
//   install  — та же стрелка вниз на поднос («установить»)
//
// Дуга при morph выпрямляется в стержень стрелки, поднос
// вырастает из точки начала дуги; стрелка при переходе в ring
// сворачивается в кольцо; кольцо к готовности разворачивается
// обратно в стрелку. Всё это рисуется каждый кадр по текущему
// morph (анимируется снаружи пружиной) и progress.
// ============================================================
Canvas {
    id: root

    property string modeA: "refresh"   // из какой формы
    property string modeB: "refresh"   // в какую форму
    property real morph: 0             // 0..1 (анимируется снаружи)
    property real progress: 0          // заполнение кольца 0..1
    property color iconColor: "#000000"
    property real strokeW: 2.4

    readonly property real m: morph < 0 ? 0 : (morph > 1 ? 1 : morph)

    onMChanged: requestPaint()
    onModeAChanged: requestPaint()
    onModeBChanged: requestPaint()
    onProgressChanged: requestPaint()
    onIconColorChanged: requestPaint()
    onPaint: draw()

    function lerp(a, b, t) { return a + (b - a) * t }
    function clamp01(v) { return v < 0 ? 0 : (v > 1 ? 1 : v) }

    // Точки форм в 24-единичном пространстве
    function shapePts(mode) {
        var shaft = [], head = [], extra = []
        var i, a
        if (mode === "refresh") {
            // дуга от -50° до 200° (по часовой), r=7.6, центр (12,12)
            for (i = 0; i < 11; ++i) {
                a = (-50 + 250 * i / 10) * Math.PI / 180
                shaft.push([12 + 7.6 * Math.cos(a), 12 + 7.6 * Math.sin(a)])
            }
            var eA = 200 * Math.PI / 180
            var ex = 12 + 7.6 * Math.cos(eA), ey = 12 + 7.6 * Math.sin(eA)
            var tx = -Math.sin(eA), ty = Math.cos(eA)   // касательная (по часовой)
            var nx = Math.cos(eA), ny = Math.sin(eA)    // нормаль наружу
            head.push([ex - tx * 3.4 + nx * 2.8, ey - ty * 3.4 + ny * 2.8])
            head.push([ex, ey])
            head.push([ex - tx * 3.4 - nx * 2.8, ey - ty * 3.4 - ny * 2.8])
            var sA = -50 * Math.PI / 180
            var sx = 12 + 7.6 * Math.cos(sA), sy = 12 + 7.6 * Math.sin(sA)
            for (i = 0; i < 4; ++i) extra.push([sx, sy])
        } else if (mode === "download" || mode === "install") {
            for (i = 0; i < 11; ++i)
                shaft.push([12, lerp(4.6, 13.0, i / 10)])
            head.push([7.4, 9.4])
            head.push([12, 14.6])
            head.push([16.6, 9.4])
            for (i = 0; i < 4; ++i)
                extra.push([lerp(6.6, 17.4, i / 3), 19.6])
        } else { // ring — кольцо прогресса
            var sweep = Math.PI * 2 * clamp01(progress)
            for (i = 0; i < 11; ++i) {
                a = -Math.PI / 2 + sweep * i / 10
                shaft.push([12 + 8.0 * Math.cos(a), 12 + 8.0 * Math.sin(a)])
            }
            for (i = 0; i < 3; ++i) head.push([12, 4.0])
            for (i = 0; i < 4; ++i) extra.push([12, 4.0])
        }
        return { shaft: shaft, head: head, extra: extra }
    }

    function strokePolyline(ctx, pts) {
        if (pts.length < 2) return
        // вырожденный штрих (все точки в одну) не рисуем
        var len = 0
        for (var i = 1; i < pts.length; ++i)
            len += Math.abs(pts[i][0] - pts[i-1][0]) + Math.abs(pts[i][1] - pts[i-1][1])
        if (len < 0.6) return
        ctx.beginPath()
        ctx.moveTo(pts[0][0], pts[0][1])
        for (i = 1; i < pts.length; ++i)
            ctx.lineTo(pts[i][0], pts[i][1])
        ctx.stroke()
    }

    function draw() {
        var ctx = getContext("2d")
        ctx.reset()
        ctx.scale(width / 24, height / 24)
        var t = m
        var A = shapePts(modeA), B = shapePts(modeB)

        // дорожка кольца (пока форма ring участвует в интерполяции)
        var ringW = (modeA === "ring" ? 1 - t : 0) + (modeB === "ring" ? t : 0)
        if (ringW > 0.01) {
            ctx.lineWidth = strokeW
            ctx.lineCap = "round"
            ctx.strokeStyle = Qt.rgba(iconColor.r, iconColor.g, iconColor.b, 0.25 * ringW)
            ctx.beginPath()
            ctx.arc(12, 12, 8.0, 0, Math.PI * 2)
            ctx.stroke()
        }

        ctx.strokeStyle = iconColor
        ctx.lineWidth = strokeW
        ctx.lineCap = "round"
        ctx.lineJoin = "round"
        var groups = ["shaft", "head", "extra"]
        for (var g = 0; g < 3; ++g) {
            var pa = A[groups[g]], pb = B[groups[g]]
            if (pa.length !== pb.length) continue
            var pts = []
            for (var i = 0; i < pa.length; ++i)
                pts.push([lerp(pa[i][0], pb[i][0], t), lerp(pa[i][1], pb[i][1], t)])
            strokePolyline(ctx, pts)
        }
    }
}
