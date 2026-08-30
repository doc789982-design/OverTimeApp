import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

AppSidePanel {
    id: root
    width: 550
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

    // Небольшая плашка с QR-кодом (как было раньше), но картинка — высокого
    // разрешения (icons/github_qr.png). Номер версии прижат к самому низу окна.
    Item {
        width: parent.width
        height: (parent && parent.parent && parent.parent.availableHeight)
                ? parent.parent.availableHeight : root.height

        Column {
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.top: parent.top
            anchors.topMargin: AppTheme.spaceXL
            anchors.bottom: versionText.top
            anchors.bottomMargin: AppTheme.spaceL
            spacing: AppTheme.spaceL

            // ==========================================
            // ПЛАШКА С QR-КОДОМ
            // ==========================================
            Rectangle {
                width: parent.width
                height: githubRow.implicitHeight + (AppTheme.spaceM * 2)
                color: githubHov.containsMouse ? AppTheme.bgElevated : AppTheme.bgSurface
                radius: AppTheme.radiusMedium
                border.color: AppTheme.borderDivider
                border.width: 1
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                Row {
                    id: githubRow
                    anchors.left: parent.left
                    anchors.right: parent.right
                    anchors.verticalCenter: parent.verticalCenter
                    anchors.margins: AppTheme.spaceM
                    spacing: AppTheme.spaceM

                    Rectangle {
                        width: 128
                        height: 128
                        color: "#FFFFFF"
                        radius: AppTheme.radiusSmall
                        border.color: AppTheme.borderDivider
                        border.width: 1

                        Image {
                            anchors.fill: parent
                            anchors.margins: 8
                            source: "../icons/github_qr.png"
                            fillMode: Image.PreserveAspectFit
                            smooth: true
                            mipmap: true
                            asynchronous: true
                        }
                    }

                    Column {
                        width: parent.width - 128 - AppTheme.spaceM
                        spacing: AppTheme.spaceXS
                        y: (parent.height - height) / 2

                        Text {
                            width: parent.width
                            text: "Скачать актуальную версию"
                            color: AppTheme.accentBrand
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBodyLarge
                            font.weight: AppTheme.weightBold
                            wrapMode: Text.WordWrap
                        }
                        Text {
                            width: parent.width
                            text: "Актуальную версию программы всегда можно скачать по ссылке на GitHub."
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            wrapMode: Text.WordWrap
                        }
                        Text {
                            width: parent.width
                            text: "github.com/doc789982-design/OverTimeApp/releases"
                            color: AppTheme.accentBrand
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightMedium
                            wrapMode: Text.WrapAnywhere
                        }
                    }
                }

                MouseArea {
                    id: githubHov
                    anchors.fill: parent
                    cursorShape: Qt.PointingHandCursor
                    hoverEnabled: true
                    onClicked: Qt.openUrlExternally(root.githubUrl)
                }
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
