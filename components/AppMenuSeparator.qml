import QtQuick
import QtQuick.Controls

MenuSeparator {
    contentItem: Rectangle {
        implicitWidth: 180
        implicitHeight: 1
        color: AppTheme.borderDivider // Строго системный цвет линии
        anchors.horizontalCenter: parent.horizontalCenter
    }
}