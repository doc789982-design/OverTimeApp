import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Item {
    id: root

    property string title: ""
    property string description: ""
    default property alias content: contentColumn.data

    // Страница прокручивается, чтобы при маленьком окне настроек
    // (которое масштабируется вместе с главным окном) контент не обрезался.
    ScrollView {
        id: pageScroll
        anchors.fill: parent
        clip: true
        ScrollBar.vertical.policy: ScrollBar.AsNeeded
        ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

        ColumnLayout {
            // Единые отступы слева/справа, как в остальных окнах
            x: AppTheme.spaceXL
            width: Math.max(0, pageScroll.availableWidth - AppTheme.spaceXL * 2)
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
            }
        }
    }
}
