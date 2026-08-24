import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

// ==========================================
// УНИВЕРСАЛЬНЫЙ ДИАЛОГ ПОДТВЕРЖДЕНИЯ
// Стиль: как у современных приложений —
// затемнение фона, иконка, внятный вопрос,
// опасная кнопка выделена красным.
// ==========================================
Popup {
    id: root

    property string dialogTitle: "Вы уверены?"
    property string dialogMessage: ""
    property string confirmLabel: "Подтвердить"
    property bool dangerMode: true   // true = красная кнопка (удаление и т.п.)

    signal accepted()
    signal rejected()

    anchors.centerIn: parent

    width: 420
    height: contentColumn.implicitHeight + AppTheme.spaceL * 2 + 60

    modal: true
    dim: true
    focus: true
    closePolicy: Popup.CloseOnEscape   // Клик мимо окна НЕ закрывает (защита от случайностей)

    z: AppTheme.zModal

    // Затемнение фона — как у больших модальных окон программы
    Overlay.modal: Rectangle {
        color: AppTheme.bgOverlay
        opacity: root.opened ? 1.0 : 0.0
        Behavior on opacity {
            NumberAnimation { duration: AppTheme.durStandard; easing.type: root.opened ? AppTheme.easeEnter : AppTheme.easeExit }
        }
    }

    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: 150; easing.type: Easing.OutQuad }
            NumberAnimation { property: "scale"; from: 0.9; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: 100; easing.type: Easing.InQuad }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: 150; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        color: AppTheme.bgModal
        radius: AppTheme.radiusModal
        border.color: AppTheme.borderDivider
        border.width: 1
        AppShadow { level: 4 }
    }

    contentItem: Item {
        Column {
            id: contentColumn
            x: AppTheme.spaceL
            width: root.width - AppTheme.spaceL * 2
            spacing: AppTheme.spaceM

            Item { width: 1; height: AppTheme.spaceS }

            // ИКОНКА В КРУЖКЕ
            Rectangle {
                width: 44
                height: 44
                radius: AppTheme.radiusPill
                anchors.horizontalCenter: parent.horizontalCenter
                color: root.dangerMode ? AppTheme.bgDangerSoft : AppTheme.bgBrandSoft

                IconImage {
                    anchors.centerIn: parent
                    source: root.dangerMode ? "../icons/trash.svg" : "../icons/help.svg"
                    width: AppTheme.iconLarge
                    height: AppTheme.iconLarge
                    color: root.dangerMode ? AppTheme.accentDanger : AppTheme.accentBrand
                }
            }

            // ЗАГОЛОВОК
            Text {
                width: parent.width
                text: root.dialogTitle
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH3
                font.weight: AppTheme.weightBold
                horizontalAlignment: Text.AlignHCenter
                wrapMode: Text.WordWrap
            }

            // ОПИСАНИЕ
            Text {
                width: parent.width
                text: root.dialogMessage
                visible: root.dialogMessage !== ""
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                horizontalAlignment: Text.AlignHCenter
                wrapMode: Text.WordWrap
                lineHeight: 1.25
            }
        }

        // КНОПКИ
        Row {
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.bottom: parent.bottom
            anchors.bottomMargin: AppTheme.spaceL
            spacing: AppTheme.spaceM

            AppButton {
                text: "Отмена"
                variant: "secondary"
                onClicked: { root.rejected(); root.close() }
            }

            AppButton {
                text: root.confirmLabel
                variant: root.dangerMode ? "danger" : "primary"
                onClicked: { root.accepted(); root.close() }
            }
        }
    }

    function showCentered() {
        root.open()
    }
}
