import QtQuick
import QtQuick.Controls.impl

Item {
    id: root
    anchors.fill: parent
    clip: true

    // Серый цвет — один, тихий, ненавязчивый
    property color sparkColor: AppTheme.textTertiary

    component Spark: Item {
        id: sparkRef
        width: 14
        height: 14

        property int startX: 0
        property int startY: 0
        property int delay: 0
        property int duration: 5000

        x: startX
        y: startY
        opacity: 0.0
        scale: 0.5

        IconImage {
            anchors.fill: parent
            source: "../icons/sparkle.svg"
            color: root.sparkColor
        }

        SequentialAnimation {
            loops: Animation.Infinite
            running: root.visible

            PauseAnimation { duration: sparkRef.delay }

            ParallelAnimation {
                NumberAnimation {
                    target: sparkRef
                    property: "y"
                    from: startY
                    to: startY - 15
                    duration: sparkRef.duration
                    easing.type: Easing.OutSine
                }

                NumberAnimation {
                    target: sparkRef
                    property: "rotation"
                    from: 0
                    to: 60
                    duration: sparkRef.duration
                }

                SequentialAnimation {
                    NumberAnimation {
                        target: sparkRef
                        property: "opacity"
                        from: 0.0
                        to: 0.25
                        duration: sparkRef.duration * 0.4
                        easing.type: Easing.InOutQuad
                    }
                    NumberAnimation {
                        target: sparkRef
                        property: "opacity"
                        from: 0.25
                        to: 0.0
                        duration: sparkRef.duration * 0.6
                        easing.type: Easing.InOutQuad
                    }
                }

                SequentialAnimation {
                    NumberAnimation {
                        target: sparkRef
                        property: "scale"
                        from: 0.2
                        to: 1.0
                        duration: sparkRef.duration * 0.5
                        easing.type: Easing.OutBack
                    }
                    NumberAnimation {
                        target: sparkRef
                        property: "scale"
                        from: 1.0
                        to: 0.4
                        duration: sparkRef.duration * 0.5
                        easing.type: Easing.InSine
                    }
                }
            }
        }
    }

    // Одна звёздочка по центру ячейки
    Spark {
        startX: parent.width * 0.5 - 7
        startY: parent.height * 0.5
        delay: 0
        duration: 5000
    }
}