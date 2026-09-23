import QtQuick
import QtQuick.Controls.impl

// ═══════════════════════════════════════════════════════════════════
// ИСКРЫ-СТРЕЛОЧКИ У КНОПКИ ОБНОВЛЕНИЯ (шапка окна)
//
// Когда появилось обновление — или новая версия уже скачана — вокруг
// кнопки загрузки в шапке тихо искрят маленькие стрелочки того же
// вида, что и сама кнопка: привлекают внимание, не мигая и не
// перекрикивая интерфейс.
//
// Зона искрения намеренно небольшая (clip держит её в рамках),
// эффект живёт прямо в главном окне — никаких поверх всех окон.
//
// Анимации — те же, что у праздничных искр в календаре: всплытие,
// «дыхание» размера и прозрачности. Видеокарта не считает ничего
// тяжёлого: это только трансформации готовых картинок, и они спят,
// когда окно скрыто или обновления нет.
// ═══════════════════════════════════════════════════════════════════
Item {
    id: root

    // Зона искрения: компактная, по центру вокруг кнопки
    width: 170
    height: 76
    clip: true

    // Искрим, только когда обновление доступно (но не качается)
    // или уже скачано и готово к установке
    property bool active: false

    // Режим «скачано»: чуть живее и ярче
    readonly property bool readyMode: backend.updateReady

    visible: active
    readonly property color sparkColor: AppTheme.accentSuccess

    // ── Одна искра-стрелочка ──────────────────────────────────────
    component Spark: Item {
        id: sparkRef
        width: sparkSize
        height: sparkSize

        // Параметры внешнего вида
        property int startX: 0
        property int startY: 0
        property int sparkSize: 11
        property int delay: 0
        property int duration: 2000
        property int rise: 16
        property real peak: 0.8          // максимум прозрачности
        property int wiggle: 20          // лёгкое покачивание

        x: startX
        y: startY
        opacity: 0.0
        scale: 0.3

        IconImage {
            anchors.fill: parent
            source: "../icons/export_arrow.svg"
            color: root.sparkColor
        }

        SequentialAnimation {
            loops: Animation.Infinite
            // Спим, когда искры не нужны или окно скрыто/свёрнуто
            running: root.visible && root.active
                     && root.Window.window !== null && root.Window.window.visible

            PauseAnimation { duration: sparkRef.delay }

            ParallelAnimation {
                // Всплытие вверх
                NumberAnimation {
                    target: sparkRef; property: "y"
                    from: sparkRef.startY; to: sparkRef.startY - sparkRef.rise
                    duration: sparkRef.duration; easing.type: Easing.OutSine
                }
                // Лёгкое покачивание стрелочки
                NumberAnimation {
                    target: sparkRef; property: "rotation"
                    from: -sparkRef.wiggle; to: sparkRef.wiggle
                    duration: sparkRef.duration; easing.type: Easing.InOutSine
                }
                // Появление и затухание
                SequentialAnimation {
                    NumberAnimation {
                        target: sparkRef; property: "opacity"
                        from: 0.0; to: sparkRef.peak
                        duration: sparkRef.duration * 0.3; easing.type: Easing.OutQuad
                    }
                    NumberAnimation {
                        target: sparkRef; property: "opacity"
                        from: sparkRef.peak; to: 0.0
                        duration: sparkRef.duration * 0.7; easing.type: Easing.InQuad
                    }
                }
                // «Дыхание» размера
                SequentialAnimation {
                    NumberAnimation {
                        target: sparkRef; property: "scale"
                        from: 0.25; to: 1.0
                        duration: sparkRef.duration * 0.4; easing.type: Easing.OutBack
                    }
                    NumberAnimation {
                        target: sparkRef; property: "scale"
                        from: 1.0; to: 0.4
                        duration: sparkRef.duration * 0.6; easing.type: Easing.InSine
                    }
                }
            }
        }
    }

    // ── Рассыпание искр вокруг кнопки ─────────────────────────────
    // В режиме «скачано» искры чуть быстрее и ярче (перезапуск через
    // привязку к readyMode не нужен: параметры читаются на каждом цикле)
    Spark { startX: 10;  startY: 56; delay: 0;    duration: root.readyMode ? 1500 : 2100; rise: 20; peak: root.readyMode ? 0.95 : 0.75; sparkSize: 12; wiggle: 24 }
    Spark { startX: 38;  startY: 48; delay: 450;  duration: root.readyMode ? 1400 : 1900; rise: 24; peak: root.readyMode ? 0.9  : 0.7;  sparkSize: 10; wiggle: 30 }
    Spark { startX: 66;  startY: 60; delay: 900;  duration: root.readyMode ? 1600 : 2300; rise: 18; peak: root.readyMode ? 0.85 : 0.65; sparkSize: 13; wiggle: 18 }
    Spark { startX: 92;  startY: 50; delay: 300;  duration: root.readyMode ? 1500 : 2000; rise: 22; peak: root.readyMode ? 0.9  : 0.7;  sparkSize: 9;  wiggle: 28 }
    Spark { startX: 118; startY: 58; delay: 700;  duration: root.readyMode ? 1400 : 2200; rise: 26; peak: root.readyMode ? 0.95 : 0.75; sparkSize: 11; wiggle: 22 }
    Spark { startX: 146; startY: 48; delay: 1100; duration: root.readyMode ? 1600 : 1900; rise: 20; peak: root.readyMode ? 0.85 : 0.65; sparkSize: 10; wiggle: 26 }
}
