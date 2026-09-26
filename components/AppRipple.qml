import QtQuick

// ============================================================
// RIPPLE — ВОЛНА ОТ НАЖАТИЯ В СТИЛЕ MATERIAL DESIGN 3
//
// Точка касания берётся PointHandler'ом: он следит за указателем,
// ничего не перехватывая, поэтому клики по кнопке работают как
// обычно. Волна — круг, который вырастает из точки нажатия
// (350 мс с замедлением) и растворяется (250 мс). Пик яркости —
// как state layer MD3 (~12–16%), поэтому волна не кричит.
//
// Вставляется внутрь нажимаемого элемента:
//     Rectangle { ... AppRipple { rippleColor: ... } }
// ============================================================
Item {
    id: host

    // Цвет волны: на залитых кнопках — цвет текста кнопки,
    // на прозрачных — бренд. Альфа уже учтена в цвете.
    property color rippleColor: AppTheme.isDark ? Qt.rgba(1, 1, 1, 0.14)
                                                : Qt.rgba(0, 0, 0, 0.12)

    // Можно запретить волны точечно (например, для ghost-кнопок)
    property bool rippleEnabled: true

    anchors.fill: parent
    clip: true

    PointHandler {
        id: ph
        onActiveChanged: {
            if (active && host.rippleEnabled) {
                wave.x = ph.point.position.x - wave.width / 2
                wave.y = ph.point.position.y - wave.height / 2
                rippleAnim.restart()
            }
        }
    }

    Rectangle {
        id: wave
        objectName: "rippleWave"
        // диаметр с запасом покрывает любую кнопку от точки нажатия
        width: Math.max(host.width, host.height) * 2.2
        height: width
        radius: width / 2
        color: host.rippleColor
        opacity: 0
        scale: 0
    }

    SequentialAnimation {
        id: rippleAnim
        objectName: "rippleAnim"
        ParallelAnimation {
            NumberAnimation { target: wave; property: "scale"; from: 0; to: 1; duration: 350; easing.type: Easing.OutQuad }
            NumberAnimation { target: wave; property: "opacity"; from: 0; to: 1; duration: 120; easing.type: Easing.OutQuad }
        }
        NumberAnimation { target: wave; property: "opacity"; from: 1; to: 0; duration: 250; easing.type: Easing.InQuad }
    }
}
