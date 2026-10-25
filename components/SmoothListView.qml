import QtQuick

// ============================================================
// СПИСОК С ПЛАВНОЙ ПРОКРУТКОЙ КОЛЕСОМ — «КАК В БРАУЗЕРЕ»
//
// Обычный ListView плюс браузерное поведение колеса: щелчок
// добавляет шаг к цели, анимация едет с замедлением и на лету
// перенацеливается — серия щелчков складывается в разгон.
// Тачпад (пиксельные дельты) двигает контент сразу, без
// анимации. Логика та же, что в SmoothFlickable.qml — держим
// файлы синхронными.
// ============================================================
ListView {
    id: root

    // Сколько пикселей проезжает один щелчок колеса (как в браузерах)
    property real wheelStep: 110
    // Идёт ли сейчас колесная анимация (прокрутка «по инерции»)
    readonly property bool wheelScrolling: wheelAnim.running
    // Накопленная цель плавной прокрутки (растёт от серии щелчков)
    property real wheelTargetY: contentY

    WheelHandler {
        acceptedDevices: PointerDevice.Mouse | PointerDevice.TouchPad
        onWheel: (ev) => {
            if (ev.pixelDelta.x !== 0 || ev.pixelDelta.y !== 0) {
                // Тачпад: точные высокочастотные события — без анимации
                wheelAnim.stop()
                root.contentY = root.clampScrollY(root.contentY - ev.pixelDelta.y)
                root.wheelTargetY = root.contentY
            } else if (ev.angleDelta.y !== 0) {
                root.wheelScrollBy(-ev.angleDelta.y * (root.wheelStep / 120))
            } else {
                ev.accepted = false // например, только горизонтальная дельта
            }
        }
    }

    NumberAnimation {
        id: wheelAnim
        target: root
        property: "contentY"
        easing.type: Easing.OutCubic
    }

    function clampScrollY(y) {
        var maxY = Math.max(originY, contentHeight - height)
        return Math.max(originY, Math.min(y, maxY))
    }

    function wheelScrollBy(dy) {
        var from = wheelAnim.running ? root.wheelTargetY : root.contentY
        var to = clampScrollY(from + dy)
        if (to === from) {
            // упёрлись в край — гасим накопленный «заряд»
            wheelAnim.stop()
            root.wheelTargetY = root.contentY
            return
        }
        wheelAnim.duration = Math.min(650, 140 + Math.abs(to - root.contentY) * 0.9)
        wheelAnim.to = to
        root.wheelTargetY = to
        wheelAnim.restart()
    }

    function cancelWheelScroll() {
        wheelAnim.stop()
        wheelTargetY = contentY
    }

    // Пользователь потащил контент мышью или отпустил «флик» —
    // колесной анимации больше не место. От самой анимации эти
    // сигналы не срабатывают (проверено).
    onMovementStarted: wheelAnim.stop()
    onFlickStarted: wheelAnim.stop()
}
