import QtQuick
import QtQuick.Controls.impl

Item {
    id: root
    anchors.fill: parent
    clip: true // Запрещаем искрам вылетать за пределы ячейки

    // Используем золотой (Warning) или фиолетовый (Purple) цвет из темы
    property color sparkColor: AppTheme.accentWarning 

    // Компонент одной искры
    component Spark: Item {
        id: sparkRef
        width: 16
        height: 16
        
        property int startX: 0
        property int startY: 0
        property int delay: 0
        property int duration: 4000 // Очень медленно (4 секунды)

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
                // Плавное всплытие наверх
                NumberAnimation { target: sparkRef; property: "y"; from: startY; to: startY - 20; duration: sparkRef.duration; easing.type: Easing.OutSine }
                
                // Медленное вращение
                NumberAnimation { target: sparkRef; property: "rotation"; from: 0; to: 90; duration: sparkRef.duration }
                
                // Дыхание (появление и затухание)
                SequentialAnimation {
                    NumberAnimation { target: sparkRef; property: "opacity"; from: 0.0; to: 0.4; duration: sparkRef.duration * 0.4; easing.type: Easing.InOutQuad }
                    NumberAnimation { target: sparkRef; property: "opacity"; from: 0.4; to: 0.0; duration: sparkRef.duration * 0.6; easing.type: Easing.InOutQuad }
                }
                
                // Дыхание размера
                SequentialAnimation {
                    NumberAnimation { target: sparkRef; property: "scale"; from: 0.2; to: 1.2; duration: sparkRef.duration * 0.5; easing.type: Easing.OutBack }
                    NumberAnimation { target: sparkRef; property: "scale"; from: 1.2; to: 0.5; duration: sparkRef.duration * 0.5; easing.type: Easing.InSine }
                }
            }
        }
    }

    // Расставляем 4 искры в разных частях ячейки с разным временем запуска
    Spark { startX: parent.width * 0.1; startY: parent.height * 0.7; delay: 0; duration: 3500 }
    Spark { startX: parent.width * 0.7; startY: parent.height * 0.8; delay: 1200; duration: 4200 }
    Spark { startX: parent.width * 0.4; startY: parent.height * 0.4; delay: 2500; duration: 3800 }
    Spark { startX: parent.width * 0.8; startY: parent.height * 0.3; delay: 3500; duration: 4500; scale: 0.8 }
}