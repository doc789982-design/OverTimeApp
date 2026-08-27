import QtQuick
import QtQuick.Controls.impl

// ============================================================
// СТРОКА ПУНКТА МЕНЮ (для морфинг-окна дня)
// ============================================================
Item {
    id: row

    property string text: ""
    property string iconSource: ""
    property string customColor: ""
    property bool showDelete: false
    signal clicked()
    signal deleteClicked()

    width: parent.width
    implicitHeight: 36

    Rectangle {
        anchors.fill: parent
        radius: AppTheme.radiusSmall
        color: rowHover.pressed ? AppTheme.statePress
              : (rowHover.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    Row {
        anchors.left: parent.left
        anchors.leftMargin: AppTheme.spaceS
        anchors.verticalCenter: parent.verticalCenter
        spacing: AppTheme.spaceS
        IconImage {
            anchors.verticalCenter: parent.verticalCenter
            source: row.iconSource
            width: AppTheme.iconMedium
            height: AppTheme.iconMedium
            color: row.customColor !== "" ? row.customColor : AppTheme.textSecondary
        }
        Text {
            anchors.verticalCenter: parent.verticalCenter
            text: row.text
            color: row.customColor !== "" ? row.customColor : AppTheme.textPrimary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody
        }
    }

    MouseArea {
        id: rowHover
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        onClicked: (mouse) => {
            if (row.showDelete && delMouse.containsMouse) return
            row.clicked()
        }
    }

    // Кнопка удаления поверх строки, чтобы клик не перехватывала строка
    Rectangle {
        id: delBtn
        visible: row.showDelete
        width: 26; height: 26; radius: AppTheme.radiusSmall
        anchors.right: parent.right
        anchors.rightMargin: 6
        anchors.verticalCenter: parent.verticalCenter
        color: delMouse.containsMouse ? AppTheme.bgDangerSoft : "transparent"
        IconImage {
            anchors.centerIn: parent
            source: "../icons/trash.svg"
            width: AppTheme.iconSmall
            height: AppTheme.iconSmall
            color: delMouse.containsMouse ? AppTheme.accentDanger : AppTheme.textTertiary
        }
        MouseArea {
            id: delMouse
            anchors.fill: parent
            hoverEnabled: true
            preventStealing: true
            cursorShape: Qt.PointingHandCursor
            onPressed: (mouse) => { mouse.accepted = true }
            onClicked: (mouse) => { row.deleteClicked(); mouse.accepted = true }
        }
    }
}
