import QtQuick
import QtQuick.Controls

Item {
    id: root

    property string title: ""
    property string description: ""
    default property alias content: contentColumn.data

    // Вертикальная прокрутка страницы. Горизонтальной не бывает:
    // ширина колонки жёстко привязана к видимой области.
    Flickable {
        id: pageScroll
        anchors.fill: parent
        clip: true
        contentWidth: width
        contentHeight: pageColumn.y + pageColumn.implicitHeight + AppTheme.spaceXL
        boundsBehavior: Flickable.StopAtBounds

        ScrollBar.vertical: ScrollBar {
            policy: ScrollBar.AsNeeded
        }

        Column {
            id: pageColumn
            x: AppTheme.spaceXL
            y: AppTheme.spaceXL
            width: pageScroll.width - AppTheme.spaceXL * 2
            spacing: AppTheme.spaceL

            Text {
                text: root.title
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH2
                font.weight: AppTheme.weightBold
            }
            Text {
                text: root.description
                width: parent.width
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                wrapMode: Text.WordWrap
            }

            Column {
                id: contentColumn
                width: parent.width
                spacing: AppTheme.spaceL
            }
        }
    }
}
