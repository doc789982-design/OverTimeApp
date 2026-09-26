import QtQuick

// ============================================================
// RIPPLE — ВОЛНА ОТ НАЖАТИЯ (MD3, адаптировано под мышь)
//
// Точка касания берётся PointHandler'ом: он следит за указателем,
// ничего не перехватывая, поэтому клики работают как обычно.
//
// ФОРМА ВОЛНЫ:
//   rippleShape: 0 (круг, по умолчанию) — круг из точки нажатия,
//     диаметр = min(width, height) − 4: волна вписана в элемент,
//     никогда не дорезается clip'ом до «квадрата» и не вылезает
//     за края — на десктопе компактная волна под курсором;
//   rippleShape: 1 (пилюля) — волна по форме самого элемента
//     (трек переключателя 52×32): растёт из центра;
//   centered: true — волна из центра элемента (чекбокс по MD3),
//     waveDiameter задаёт точный диаметр круга.
//
// ВАЖНО: доставка событий идёт сверху вниз. Если ВЫШЕ волны
// объявлена MouseArea или другой перехватчик нажатия — волна
// события не увидит. AppRipple должен быть ПОСЛЕДНИМ (самым
// верхним) ребёнком нажимаемого элемента.
//
// Тайминги — быстрее мобильного MD3: рост 220 мс, вспышка 80 мс,
// затухание 160 мс — под десктопный темп взаимодействия.
// ============================================================

Item {
    id: host

    // Цвет волны: на залитых кнопках — цвет текста кнопки,
    // на прозрачных — бренд. Альфа уже учтена в цвете.
    property color rippleColor: AppTheme.isDark ? Qt.rgba(1, 1, 1, 0.14)
                                                : Qt.rgba(0, 0, 0, 0.12)

    // Можно запретить волны точечно (например, для ghost-кнопок)
    property bool rippleEnabled: true

    property int rippleShape: 0        // 0 — круг, 1 — пилюля по форме элемента
    property bool centered: false      // волна из точки нажатия / из центра
    property real waveDiameter: -1     // диаметр круга; -1 = min(width, height) − 4

    readonly property real waveD: Math.max(12,
        waveDiameter > 0 ? waveDiameter : Math.min(width, height) - 4)

    // Откуда растёт волна. По умолчанию — центр элемента
    // (тесты гоняют волну без реального клика).
    property point waveOrigin: Qt.point(width / 2, height / 2)

    anchors.fill: parent
    clip: true

    function startRipple(cx, cy) {
        waveOrigin = (rippleShape === 1 || centered)
                     ? Qt.point(width / 2, height / 2)
                     : Qt.point(cx, cy)
        rippleAnim.restart()
    }

    PointHandler {
        id: ph
        onActiveChanged: {
            if (active && host.rippleEnabled && host.width > 0)
                host.startRipple(ph.point.position.x, ph.point.position.y)
        }
    }

    Rectangle {
        id: wave
        objectName: "rippleWave"
        // круг: вписанный в меньшую сторону; пилюля: размер элемента
        width: host.rippleShape === 1 ? host.width : host.waveD
        height: host.rippleShape === 1 ? host.height : host.waveD
        radius: height / 2
        color: host.rippleColor
        opacity: 0
        scale: 0.06
        x: host.waveOrigin.x - width / 2
        y: host.waveOrigin.y - height / 2
    }

    SequentialAnimation {
        id: rippleAnim
        objectName: "rippleAnim"
        ParallelAnimation {
            NumberAnimation { target: wave; property: "scale"; from: 0.06; to: 1; duration: 220; easing.type: Easing.OutQuad }
            NumberAnimation { target: wave; property: "opacity"; from: 0; to: 1; duration: 80; easing.type: Easing.OutQuad }
        }
        NumberAnimation { target: wave; property: "opacity"; from: 1; to: 0; duration: 160; easing.type: Easing.InQuad }
    }
}
