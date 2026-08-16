import QtQuick
import QtQuick.Layouts

Item {
    id: root

    property string title: ""
    property string description: ""
    default property alias content: contentColumn.data

    ColumnLayout {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceXL
        spacing: AppTheme.spaceL

        Column {
            spacing: AppTheme.spaceXS
            Layout.fillWidth: true

            Text {
                text: root.title
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH2
                font.weight: AppTheme.weightBold
            }
            Text {
                text: root.description
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                wrapMode: Text.WordWrap
                width: parent.width
            }
        }

        Column {
            id: contentColumn
            Layout.fillWidth: true
            spacing: AppTheme.spaceM

            // implicitHeight пробрасывается через Layout автоматически
        }

        Item {
            Layout.fillHeight: true
        }
    }
}