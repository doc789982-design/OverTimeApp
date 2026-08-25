import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

AppDialog {
    id: root
    title: "Что нового"
    acceptText: "Понятно"
    rejectText: "Закрыть"
    acceptVariant: "primary"
    width: 520
    closePolicy: Popup.CloseOnEscape

    onAccepted: root.close()
    onClosed: backend.ackWhatsNew()

    function showIfNeeded() {
        if (backend.whatsNew && backend.whatsNew.length > 0)
            root.showCentered()
    }

    Text {
        width: parent.width
        text: "После обновления. Базы и настройки на месте."
        color: AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        wrapMode: Text.WordWrap
    }

    Repeater {
        model: backend.whatsNew

        Column {
            width: root.width - AppTheme.spaceL * 2
            spacing: AppTheme.spaceS

            Text {
                width: parent.width
                text: "Версия " + modelData.version
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH5
                font.weight: AppTheme.weightBold
            }

            Column {
                width: parent.width
                spacing: AppTheme.spaceXXS
                visible: (modelData.addedText || "").length > 0

                Text {
                    text: "Добавили"
                    color: AppTheme.accentSuccess
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                }
                Text {
                    width: parent.width
                    text: modelData.addedText
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    wrapMode: Text.WordWrap
                    lineHeight: 1.25
                }
            }

            Column {
                width: parent.width
                spacing: AppTheme.spaceXXS
                visible: (modelData.fixedText || "").length > 0

                Text {
                    text: "Починили"
                    color: AppTheme.accentBrand
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                }
                Text {
                    width: parent.width
                    text: modelData.fixedText
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    wrapMode: Text.WordWrap
                    lineHeight: 1.25
                }
            }

            Column {
                width: parent.width
                spacing: AppTheme.spaceXXS
                visible: modelData.hasChanged

                Text {
                    text: "Поменяли"
                    color: AppTheme.accentWarning
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                }
                Text {
                    width: parent.width
                    text: modelData.changedText
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    wrapMode: Text.WordWrap
                    lineHeight: 1.25
                }
            }

            Column {
                width: parent.width
                spacing: AppTheme.spaceXXS
                visible: (modelData.removedText || "").length > 0

                Text {
                    text: "Удалили"
                    color: AppTheme.accentDanger
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: AppTheme.weightBold
                }
                Text {
                    width: parent.width
                    text: modelData.removedText
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    wrapMode: Text.WordWrap
                    lineHeight: 1.25
                }
            }

            Rectangle {
                visible: index < backend.whatsNew.length - 1
                width: parent.width
                height: 1
                color: AppTheme.borderDivider
            }
        }
    }
}
