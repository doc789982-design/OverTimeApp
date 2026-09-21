import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import Qt5Compat.GraphicalEffects
import QtQuick.Controls.impl

Item {
    id: root
    property bool isYearView: false

    // ════════════════════════════════════════════════════════════════
    // ОДНА ПАНЕЛЬ — ДВА НАПОЛНЕНИЯ
    // Месяц и год показываются одним и тем же элементом: шапка с
    // заголовком, кнопкой «Деньги» и графой «Всего дней», три карточки
    // показателей и поясняющая строка снизу. Различается только
    // источник данных (summ) и подписи периода.
    // ════════════════════════════════════════════════════════════════
    readonly property var summ: root.isYearView ? backend.yearSummary : backend.monthSummary
    readonly property bool isShiftMonth: !root.isYearView && backend.monthSummary.is_shift === true
    readonly property string periodText: root.isYearView
            ? (backend.currentPeriodText.split(" ")[1] + " год")
            : backend.currentPeriodText
    readonly property string endCaption: root.isYearView ? "остаток на конец года"
                                                         : "остаток на конец месяца"

    implicitHeight: backend.selectedEmployeeId !== 0 ? mainRect.height : 0

    // Склонение: 1 день / 2 дня / 5 дней
    function daysWord(n) {
        if (n % 100 >= 11 && n % 100 <= 19) return "дней"
        var d = n % 10
        if (d === 1) return "день"
        if (d >= 2 && d <= 4) return "дня"
        return "дней"
    }

    // ==========================================
    // МИНИ-ЯЧЕЙКА ПОТОКА ВНУТРИ КАРТОЧКИ
    // Значение отцентрировано по горизонтали относительно пояснения снизу.
    // ==========================================
    component MiniStat: Column {
        id: ms
        property string valText: "—"
        property string labelText: ""
        property color valColor: AppTheme.textPrimary
        spacing: 1

        Text {
            width: ms.width
            horizontalAlignment: Text.AlignHCenter
            text: ms.valText
            color: ms.valColor
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightBold
            elide: Text.ElideRight
            textFormat: Text.StyledText
        }
        Text {
            width: ms.width
            horizontalAlignment: Text.AlignHCenter
            text: ms.labelText
            color: AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeMicro
            elide: Text.ElideRight
        }
    }

    // ==========================================
    // КАРТОЧКА ОДНОГО ПОКАЗАТЕЛЯ
    // ==========================================
    component StatCard: Rectangle {
        id: card
        property string title: ""
        property string icon: ""
        property color accent: AppTheme.accentBrand
        property string endText: "—"
        property bool endNeg: false
        property string startText: "—"
        property string accText: "—"
        property string compText: "—"
        property string caption: "остаток на конец месяца"

        Layout.fillWidth: true
        Layout.fillHeight: true
        Layout.minimumWidth: 170
        Layout.preferredHeight: 170
        radius: AppTheme.radiusMedium
        color: AppTheme.bgBase
        border.color: AppTheme.borderDivider
        border.width: 1

        ColumnLayout {
            id: cardCol
            anchors.fill: parent
            anchors.leftMargin: AppTheme.spaceM
            anchors.rightMargin: AppTheme.spaceM
            anchors.topMargin: AppTheme.spaceS
            anchors.bottomMargin: AppTheme.spaceS
            spacing: AppTheme.spaceXXS

            RowLayout {
                Layout.fillWidth: true
                spacing: AppTheme.spaceS

                Rectangle {
                    width: 32; height: 32
                    radius: AppTheme.radiusSmall
                    color: Qt.rgba(card.accent.r, card.accent.g, card.accent.b, 0.14)
                    Layout.alignment: Qt.AlignVCenter
                    IconImage {
                        anchors.centerIn: parent
                        source: "../icons/" + card.icon
                        width: 18; height: 18
                        color: card.accent
                    }
                }
                Text {
                    Layout.fillWidth: true
                    Layout.alignment: Qt.AlignVCenter
                    text: card.title
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    font.weight: AppTheme.weightBold
                    elide: Text.ElideRight
                }
            }

            Text {
                Layout.fillWidth: true
                Layout.topMargin: AppTheme.spaceXS
                text: card.endText
                color: card.endNeg ? AppTheme.accentDanger : AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: 27
                font.weight: AppTheme.weightBold
                textFormat: Text.StyledText
                elide: Text.ElideRight
            }
            Text {
                Layout.fillWidth: true
                text: card.caption
                color: AppTheme.textTertiary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeMicro
            }

            // Распорка — прижимает разделитель и поток вниз карточки
            Item { Layout.fillHeight: true }

            Rectangle {
                Layout.fillWidth: true
                height: 1
                Layout.topMargin: AppTheme.spaceXS
                color: AppTheme.borderDivider
            }

            // Поток: три равные ячейки «на начало / +начислено / −компенсировано».
            // Обычный Row с явной шириной каждой ячейки — без вложенных layout-ов,
            // чтобы ячейки всегда ложились в один ряд внутри карточки.
            Row {
                id: flowRow
                Layout.fillWidth: true
                Layout.preferredHeight: 30
                Layout.topMargin: AppTheme.spaceXXS
                spacing: 0

                MiniStat { width: flowRow.width / 3; valText: card.startText; labelText: "на начало" }
                MiniStat { width: flowRow.width / 3; valText: card.accText;  labelText: "начислено"; valColor: AppTheme.accentSuccess }
                MiniStat { width: flowRow.width / 3; valText: card.compText; labelText: "компенсировано"; valColor: AppTheme.accentDanger }
            }
        }
    }

    // ==========================================
    // ГЛАВНАЯ ПАНЕЛЬ
    // ==========================================
    Rectangle {
        id: mainRect
        visible: backend.selectedEmployeeId !== 0
        anchors.top: parent.top
        anchors.left: parent.left
        anchors.right: parent.right
        height: summaryBody.implicitHeight + AppTheme.spaceM * 2
        radius: AppTheme.radiusLarge
        color: AppTheme.bgSurface
        border.color: AppTheme.borderDivider; border.width: 1

        // Тень-картинка вместо вычисляемой (Level 1)
        AppShadow { level: 1 }

        ColumnLayout {
            id: summaryBody
            anchors.top: parent.top
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.margins: AppTheme.spaceM
            spacing: AppTheme.spaceM

            // ------------------------------------------
            // ШАПКА: заголовок + период, «Деньги», «Всего дней»
            // ------------------------------------------
            RowLayout {
                Layout.fillWidth: true
                spacing: AppTheme.spaceM

                IconImage {
                    Layout.alignment: Qt.AlignVCenter
                    source: "../icons/export_box.svg"
                    width: 18; height: 18
                    color: AppTheme.textTertiary
                }
                ColumnLayout {
                    spacing: 0
                    Text {
                        text: "Балансы"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBodyLarge
                        font.weight: AppTheme.weightBold
                    }
                    Text {
                        text: root.periodText
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                    }
                }

                // Распорка: прижимает «Деньги» и «Всего дней» к правому краю строки
                Item { Layout.fillWidth: true }

                // Кнопка «Деньги» — открывает просмотр денежных компенсаций месяца
                Rectangle {
                    visible: !root.isYearView
                    height: 28
                    width: moneyPillText.implicitWidth + AppTheme.spaceL + 6
                    radius: AppTheme.radiusSmall
                    Layout.alignment: Qt.AlignVCenter

                    color: moneyPillArea.pressed ? AppTheme.statePress
                         : moneyPillArea.containsMouse ? AppTheme.bgBrandSoft
                         : "transparent"
                    border.color: moneyPillArea.containsMouse ? AppTheme.accentBrand : AppTheme.borderInput
                    border.width: 1
                    Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                    Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }

                    Row {
                        id: moneyPillText
                        anchors.centerIn: parent
                        spacing: AppTheme.spaceXXS
                        Text {
                            text: "₽"
                            color: moneyPillArea.containsMouse ? AppTheme.accentBrand : AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                        }
                        Text {
                            text: "Деньги"
                            color: moneyPillArea.containsMouse ? AppTheme.accentBrand : AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                        }
                    }
                    MouseArea {
                        id: moneyPillArea
                        anchors.fill: parent
                        hoverEnabled: true
                        cursorShape: Qt.PointingHandCursor
                        onClicked: { backend.loadMoneyComps(); moneyInspector.show() }
                    }
                    AppToolTip {
                        anchors.horizontalCenter: parent.horizontalCenter
                        anchors.bottom: parent.top; anchors.bottomMargin: AppTheme.spaceXXS
                        isVisible: moneyPillArea.containsMouse
                        text: "Посмотреть денежные компенсации"
                    }
                }

                // «Всего дней» — переработка сотрудника в днях (справа в шапке).
                // Один элемент на оба вида: число берётся из текущего набора итогов.
                // Шрифт числа — как у даты в шапке меню дня (fontCondensed, sizeH4, bold).
                RowLayout {
                    id: totalDaysStat
                    Layout.alignment: Qt.AlignVCenter
                    spacing: AppTheme.spaceXS

                    Text {
                        text: "Всего"
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        font.weight: AppTheme.weightMedium
                        Layout.alignment: Qt.AlignVCenter
                    }
                    Text {
                        text: root.summ.total_days !== undefined ? root.summ.total_days : "—"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontCondensed
                        font.pixelSize: AppTheme.sizeH4
                        font.weight: AppTheme.weightBold
                        Layout.alignment: Qt.AlignVCenter
                    }
                    Text {
                        text: root.summ.total_days !== undefined ? root.daysWord(root.summ.total_days) : ""
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        Layout.alignment: Qt.AlignVCenter
                    }
                }
            }

            // ------------------------------------------
            // ТРИ КАРТОЧКИ ПОКАЗАТЕЛЕЙ (часы / дни / сверх нормы)
            // ------------------------------------------
            RowLayout {
                id: cardsRow
                Layout.fillWidth: true
                spacing: AppTheme.spaceM

                StatCard {
                    title: "ДВО (ночные)"
                    icon: "clock.svg"
                    accent: AppTheme.accentBrand
                    caption: root.endCaption
                    endText: root.summ.end_hours || "—"
                    endNeg: root.summ.is_hours_negative === true
                    startText: root.summ.start_hours || "—"
                    accText: root.summ.acc_hours || "—"
                    compText: root.summ.comp_hours || "—"
                }
                StatCard {
                    title: "ДДО (дни)"
                    icon: "calendar.svg"
                    accent: AppTheme.accentSuccess
                    caption: root.endCaption
                    endText: root.summ.end_days || "—"
                    endNeg: root.summ.is_days_negative === true
                    startText: root.summ.start_days || "—"
                    accText: root.summ.acc_days || "—"
                    compText: root.summ.comp_days || "—"
                }
                StatCard {
                    title: "Сверх нормы"
                    icon: "overtime.svg"
                    accent: AppTheme.accentWarning
                    caption: root.endCaption
                    endText: root.summ.end_overtime || "—"
                    endNeg: root.summ.is_overtime_negative === true
                    startText: root.summ.start_overtime || "—"
                    accText: root.summ.acc_overtime || "—"
                    compText: root.summ.comp_overtime || "—"
                }
            }

            // ------------------------------------------
            // ПОЯСНЯЮЩАЯ СТРОКА (месяц): сменный график
            // «в ночь» (синий), «праздничные» (красный), норма — нейтральная.
            // ------------------------------------------
            RowLayout {
                id: shiftRow
                Layout.fillWidth: true
                visible: root.isShiftMonth
                spacing: AppTheme.spaceL

                property color nightC: Qt.rgba(AppTheme.accentBrand.r, AppTheme.accentBrand.g, AppTheme.accentBrand.b, 0.72)
                property color holC:   Qt.rgba(AppTheme.accentDanger.r, AppTheme.accentDanger.g, AppTheme.accentDanger.b, 0.75)

                Row {
                    Layout.alignment: Qt.AlignVCenter
                    spacing: AppTheme.spaceXXS
                    IconImage { source: "../icons/night.svg";    width: 13; height: 13; color: shiftRow.nightC }
                    Text {
                        text: "в ночь " + (backend.monthSummary.shift_night || "—")
                        color: shiftRow.nightC
                        font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                    }
                }
                Row {
                    Layout.alignment: Qt.AlignVCenter
                    spacing: AppTheme.spaceXXS
                    IconImage { source: "../icons/sparkle.svg"; width: 13; height: 13; color: shiftRow.holC }
                    Text {
                        text: "праздничные " + (backend.monthSummary.shift_holiday || "—")
                        color: shiftRow.holC
                        font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                    }
                }
                Row {
                    Layout.alignment: Qt.AlignVCenter
                    spacing: AppTheme.spaceXXS
                    IconImage { source: "../icons/help.svg"; width: 13; height: 13; color: AppTheme.textTertiary }
                    Text {
                        text: "норма " + (backend.monthSummary.norm_minutes || "0")
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                    }
                }
                Item { Layout.fillWidth: true }
            }

            // ------------------------------------------
            // ПОЯСНЯЮЩАЯ СТРОКА (год): статусы дней и денежные компенсации
            // Буквы-кружки повторяют обозначения матрицы года.
            // ------------------------------------------
            RowLayout {
                Layout.fillWidth: true
                visible: root.isYearView
                spacing: AppTheme.spaceM

                Repeater {
                    model: [
                        { "letter": "Б", "soft": AppTheme.bgDangerSoft,  "accent": AppTheme.accentDanger,
                          "value": backend.yearSummary.b_days || "—" },
                        { "letter": "О", "soft": AppTheme.bgWarningSoft, "accent": AppTheme.accentWarning,
                          "value": backend.yearSummary.o_days || "—" },
                        { "letter": "К", "soft": AppTheme.bgPurpleSoft,  "accent": AppTheme.accentPurple,
                          "value": backend.yearSummary.k_days || "—" }
                    ]

                    RowLayout {
                        Layout.alignment: Qt.AlignVCenter
                        spacing: AppTheme.spaceXXS

                        Rectangle {
                            Layout.alignment: Qt.AlignVCenter
                            width: 18; height: 18
                            radius: AppTheme.radiusPill
                            color: modelData.soft
                            Text {
                                anchors.centerIn: parent
                                text: modelData.letter
                                color: modelData.accent
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeMicro
                                font.weight: AppTheme.weightBold
                            }
                        }
                        Text {
                            Layout.alignment: Qt.AlignVCenter
                            text: modelData.value
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                    }
                }

                Item { Layout.fillWidth: true }

                RowLayout {
                    Layout.alignment: Qt.AlignVCenter
                    spacing: AppTheme.spaceXXS
                    IconImage {
                        Layout.alignment: Qt.AlignVCenter
                        source: "../icons/ruble.svg"
                        width: 13; height: 13
                        color: AppTheme.textTertiary
                    }
                    Text {
                        Layout.alignment: Qt.AlignVCenter
                        // длинные суммы («36 ч. (ноч.), 12 ч. (сверх), 3 дн.») не должны
                        // выдавливать строку — обрываемся многоточием
                        Layout.maximumWidth: 320
                        text: "деньгами " + (backend.yearSummary.comp_money || "—")
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        elide: Text.ElideRight
                    }
                }
            }
        }
    }
}
