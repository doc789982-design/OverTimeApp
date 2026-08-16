// ========================================
// НАЧАЛО ФАЙЛА: IconImage.qml
// ========================================

import QtQuick
import Qt5Compat.GraphicalEffects

Image {
    id: root

    // Цвет заливки иконки (по умолчанию черный)
    property color color: "#000000"

    // Источник должен быть один и тот же, но он дублируется
    // для маски и для самой иконки
    source: ""

    // Делаем саму картинку невидимой
    visible: false

    // Компонент, который реально отображается на экране
    ColorOverlay {
        anchors.fill: parent
        color: root.color // Цвет берется из свойства 'color'

        // В качестве "маски" используется та же самая картинка
        source: Image {
            source: root.source
            sourceSize: Qt.size(root.width, root.height)
        }
    }
}