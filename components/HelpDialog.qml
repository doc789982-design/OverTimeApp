import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

AppSidePanel {
    id: root
    width: 680
    title: "О программе"

    readonly property string githubUrl: "https://github.com/doc789982-design/OverTimeApp/releases"

    // Человекочитаемое описание действия горячей клавиши (используется в подсказке
    // у иконки справки в главном окне — main.qml).
    function describeHotkey(hk) {
        if (hk.name && hk.name !== "") return hk.name

        function getPlural(n, f1, f2, f5) {
            let n10 = Math.abs(n) % 10, n100 = Math.abs(n) % 100
            if (n100 >= 11 && n100 <= 14) return f5
            if (n10 === 1) return f1
            if (n10 >= 2 && n10 <= 4) return f2
            return f5
        }

        if (hk.type === "duty") {
            let sType = hk.duty_shift ? "(в смене)" : "(вне графика)"
            let res = "Дежурство " + sType + " с " + hk.duty_start + " до " + hk.duty_end
            if (hk.duty_breaks && hk.duty_breaks.length > 0) {
                let b = hk.duty_breaks.map(item => "с " + item.start + " до " + item.end).join(", ")
                res += " с перерывом " + b
            }
            return res
        }
        if (hk.type === "comp") {
            let unit = hk.comp_unit === "days" ? getPlural(hk.comp_amount, "день", "дня", "дней")
                                               : getPlural(hk.comp_amount, "час", "часа", "часов")
            return "Компенсация " + hk.comp_amount + " " + unit
        }
        if (hk.type === "status") {
            let sNames = {"Б": "Больничный", "О": "Отпуск", "К": "Командировка"}
            return sNames[hk.status_val] || hk.status_val
        }
        return ""
    }

    // Контент занимает всю высоту панели (без прокрутки): вверху большой QR-код,
    // под ним текст, а номер версии прижат к самому низу окна.
    Item {
        width: parent.width
        height: (parent && parent.parent && parent.parent.availableHeight)
                ? parent.parent.availableHeight : root.height

        // ==========================================
        // БОЛЬШОЙ QR-КОД НА ВСЁ ОКНО СПРАВКИ
        // ==========================================
        Rectangle {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: parent.top
            anchors.bottom: textBlock.top
            anchors.bottomMargin: AppTheme.spaceL

            color: "#FFFFFF"
            radius: AppTheme.radiusLarge
            border.color: AppTheme.borderDivider
            border.width: 1

            Rectangle {
                anchors.fill: parent
                radius: parent.radius
                color: qrHov.pressed
                       ? AppTheme.statePress
                       : (qrHov.containsMouse ? Qt.rgba(0, 0, 0, 0.03) : "transparent")
            }

            Image {
                anchors.fill: parent
                anchors.margins: Math.max(AppTheme.spaceM, parent.height * 0.06)
                source: "../icons/github_qr.png"
                fillMode: Image.PreserveAspectFit
                smooth: true
                mipmap: true
                asynchronous: true
            }

            MouseArea {
                id: qrHov
                anchors.fill: parent
                cursorShape: Qt.PointingHandCursor
                hoverEnabled: true
                onClicked: Qt.openUrlExternally(root.githubUrl)
            }
        }

        // ==========================================
        // ТЕКСТ ПОД QR-КОДОМ
        // ==========================================
        Column {
            id: textBlock
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.bottom: versionText.top
            anchors.bottomMargin: AppTheme.spaceL
            spacing: AppTheme.spaceXS

            Text {
                width: parent.width
                text: "Скачать актуальную версию"
                color: AppTheme.accentBrand
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBodyLarge
                font.weight: AppTheme.weightBold
                horizontalAlignment: Text.AlignHCenter
            }
            Text {
                width: parent.width
                text: "Отсканируйте QR-код или нажмите на него, чтобы открыть страницу загрузки актуальной версии программы на GitHub."
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                horizontalAlignment: Text.AlignHCenter
                wrapMode: Text.WordWrap
            }
            Text {
                width: parent.width
                text: "github.com/doc789982-design/OverTimeApp/releases"
                color: AppTheme.accentBrand
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightMedium
                horizontalAlignment: Text.AlignHCenter
                wrapMode: Text.WrapAnywhere
            }
        }

        // ==========================================
        // НОМЕР ВЕРСИИ — САМЫЙ НИЗ ОКНА
        // ==========================================
        Text {
            id: versionText
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.bottom: parent.bottom
            text: "OVERTIMETAB · " + AppTheme.appVersionFull
            color: AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightMedium
            horizontalAlignment: Text.AlignHCenter
        }
    }
}
