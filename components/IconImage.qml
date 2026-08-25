//
// ИСТОРИЯ БАГА: раньше корнем здесь был Image { visible: false },
// а внутри него лежал ColorOverlay. В QML visible:false скрывает
// и всех детей, поэтому ColorOverlay никогда не отрисовывался —
// иконка была невидимой. Это ломало иконки (корзину и др.) в тех
// файлах, где 'import "."' перекрывал системный IconImage из
// QtQuick.Controls.impl этим локальным файлом
// (CalendarWorkspace.qml, EmployeeListPanel.qml).
//
// ТЕПЕРЬ: это тонкая обёртка над тем же системным IconImage,
// который используется во всех остальных файлах проекта.
// Поведение везде становится одинаковым.

import QtQuick
import QtQuick.Controls.impl as Impl

Impl.IconImage {
    // Совместимость со старым API: свойство 'color' задаёт цвет иконки.
    // У системного IconImage цвет тоже называется 'color', так что
    // все существующие вызовы работают без изменений.
    fillMode: Image.PreserveAspectFit
    sourceSize: Qt.size(width, height)
}
