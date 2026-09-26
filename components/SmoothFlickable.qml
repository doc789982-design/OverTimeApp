import QtQuick

// ============================================================
// ПЛАВНАЯ ПРОКРУТКА КОЛЕСОМ — «КАК В БРАУЗЕРЕ»
//
// Щелчок колеса не дёргает контент, а добавляет шаг к цели:
// анимация едет с плавным замедлением и на лету перенацеливается,
// поэтому серия щелчков складывается в разгон (чем длиннее
// оставшийся путь — тем дольше едем). Тачпад шлёт пиксельные
// дельты — их применяем сразу, без анимации, чтобы не отставать
// от пальца. Потянули контент мышью или дёрнули скроллбар —
// колесная анимация уступает дорогу.
// ============================================================
Flickable {
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
