import QtQuick
import QtQuick.Controls

MenuSeparator {
    padding: 0
    topPadding: 4
    bottomPadding: 4
    leftPadding: 12
    rightPadding: 12

    contentItem: Rectangle {
        implicitWidth: 1
        implicitHeight: 1
        color: AppTheme.borderDivider
    }
}
