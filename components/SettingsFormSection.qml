import QtQuick
import QtQuick.Layouts

Item {
    id: root

    property string title: ""
    default property alias content: innerColumn.data

    // Секции живут только внутри ColumnLayout панелей —
    // растягиваются на всю ширину панели
    Layout.fillWidth: true
    implicitHeight: innerColumn.implicitHeight
    height: implicitHeight

    Column {
        id: innerColumn
        width: parent.width
        spacing: AppTheme.spaceM

        Text {
            text: root.title
            color: AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightBold
            font.letterSpacing: 0.8
        }
    }
}