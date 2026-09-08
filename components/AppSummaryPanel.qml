import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import Qt5Compat.GraphicalEffects
import QtQuick.Controls.impl

Item {
    id: root
    property bool isYearView: false

    implicitHeight: backend.selectedEmployeeId !== 0 ? mainRect.height : 0

    readonly property int layoutMode: root.width > 900 ? 2 : (root.width > 600 ? 1 : 0)

    onLayoutModeChanged: {
        col1Wrapper.resetAndEnter(0)
        col2Wrapper.resetAndEnter(50)
        col3Wrapper.resetAndEnter(100)
        col4Wrapper.resetAndEnter(150)
    }

    // ==========================================
    // МЕХАНИЧЕСКИЙ БАРАБАН
    // ==========================================
    component OdometerValue: Item {
        id: odoRoot
        property string text: "—"
        property color textColor: AppTheme.textPrimary
        property string _lastText: "—"
        clip: true

        function extractVal(s) {
            if (!s || s === "—") return 0
            let m = s.match(/-?\d+/)
            return m ? parseInt(m[0], 10) : 0
        }

        onTextChanged: {
            if (text === _lastText) return
            anim.stop()
            let oldV = extractVal(_lastText)
            let newV = extractVal(text)
            let dir = (newV >= oldV) ? 1 : -1
            mainText.text = _lastText
            newText.text = text
            mainText.y = 0; mainText.opacity = 1.0
            newText.y = (dir === 1) ? odoRoot.height : -odoRoot.height
            newText.opacity = 0.0
            moveOld.to = (dir === 1) ? -odoRoot.height : odoRoot.height
            moveNew.from = (dir === 1) ? odoRoot.height : -odoRoot.height
            anim.start()
            _lastText = text
        }

        Text {
            id: mainText
            anchors.left: parent.left; anchors.right: parent.right
            height: parent.height
            horizontalAlignment: Text.AlignRight; verticalAlignment: Text.AlignVCenter
            font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody
            font.weight: AppTheme.weightBold
            color: odoRoot.textColor; textFormat: Text.StyledText
            text: odoRoot._lastText
        }
        Text {
            id: newText
            anchors.left: parent.left; anchors.right: parent.right
            height: parent.height
            horizontalAlignment: Text.AlignRight; verticalAlignment: Text.AlignVCenter
            font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody
            font.weight: AppTheme.weightBold
            color: odoRoot.textColor; textFormat: Text.StyledText
            opacity: 0.0
        }
        ParallelAnimation {
            id: anim
            NumberAnimation { id: moveOld; target: mainText; property: "y"; duration: 250; easing.type: Easing.OutQuart }
            NumberAnimation { target: mainText; property: "opacity"; to: 0.0; duration: 150 }
            NumberAnimation { id: moveNew; target: newText; property: "y"; to: 0; duration: 250; easing.type: Easing.OutQuart }
            NumberAnimation { target: newText; property: "opacity"; to: 1.0; duration: 150 }
        }
    }

    // ==========================================
    // СТРОКА ИТОГОВ
    // ==========================================
    component SummaryRow: RowLayout {
        id: summaryRowRoot
        property string labelText: ""
        property string valText: ""
        property color valColor: AppTheme.textPrimary

        Layout.fillWidth: true
        implicitHeight: 18
        spacing: AppTheme.spaceXS

        Text {
            Layout.fillWidth: true
            Layout.alignment: Qt.AlignVCenter
            text: summaryRowRoot.labelText
            color: AppTheme.textSecondary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            elide: Text.ElideRight
        }
        OdometerValue {
            Layout.preferredWidth: 90
            Layout.minimumWidth: 60
            implicitHeight: summaryRowRoot.implicitHeight
            text: summaryRowRoot.valText
            textColor: summaryRowRoot.valColor
        }
    }

    // ==========================================
    // ОБЁРТКА КОЛОНКИ С MD3-АНИМАЦИЕЙ
    // ==========================================
    component ColWrapper: Item {
        id: wrapperRoot
        clip: true

        function resetAndEnter(delayMs) {
            staggerTimer.stop()
            enterAnim.stop()
            wrapperRoot.opacity = 0.0
            wrapperRoot.scale = 0.94
            staggerTimer.interval = delayMs
            staggerTimer.restart()
        }

        Component.onCompleted: {
            wrapperRoot.opacity = 0.0
            wrapperRoot.scale = 0.94
            staggerTimer.interval = 0
            staggerTimer.restart()
        }

        Timer {
            id: staggerTimer
            repeat: false
            onTriggered: enterAnim.start()
        }

        ParallelAnimation {
            id: enterAnim

            NumberAnimation {
                target: wrapperRoot
                property: "opacity"
                from: 0.0; to: 1.0
                duration: 300
                easing.type: Easing.OutCubic
            }

            NumberAnimation {
                target: wrapperRoot
                property: "scale"
                from: 0.94; to: 1.0
                duration: 350
                easing.type: Easing.BezierSpline
                easing.bezierCurve: [0.05, 0.7, 0.1, 1.0, 1.0, 1.0]
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
        height: (root.isYearView ? contentArea.implicitHeight : monthBody.implicitHeight)
                + AppTheme.spaceM * 2
        radius: AppTheme.radiusLarge
        color: AppTheme.bgSurface
        border.color: AppTheme.borderDivider; border.width: 1

        // Тень-картинка вместо вычисляемой (Level 1)
        AppShadow { level: 1 }

        Item {
            id: contentArea
            visible: root.isYearView
            anchors.top: parent.top
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.margins: AppTheme.spaceM

            implicitHeight: {
                switch(root.layoutMode) {
                    case 2:
                        return Math.max(col1Content.implicitHeight,
                                        col2Content.implicitHeight,
                                        col3Content.implicitHeight,
                                        col4Content.implicitHeight)
                    case 1:
                        return Math.max(col1Content.implicitHeight, col2Content.implicitHeight)
                             + AppTheme.spaceL
                             + Math.max(col3Content.implicitHeight, col4Content.implicitHeight)
                    default:
                        return col1Content.implicitHeight + AppTheme.spaceL
                             + col2Content.implicitHeight + AppTheme.spaceL
                             + col3Content.implicitHeight + AppTheme.spaceL
                             + col4Content.implicitHeight
                }
            }

            readonly property real colW1:  contentArea.width
            readonly property real colW2: (contentArea.width - AppTheme.spaceXL) / 2
            readonly property real colW4: (contentArea.width - AppTheme.spaceXL * 3) / 4
            readonly property real row2Y: {
                if (root.layoutMode === 1)
                    return Math.max(col1Content.implicitHeight,
                                    col2Content.implicitHeight) + AppTheme.spaceL
                return 0
            }

            // ------------------------------------------
            // КОЛОНКА 1
            // ------------------------------------------
            ColWrapper {
                id: col1Wrapper
                x: 0
                y: 0
                width:  root.layoutMode === 0 ? contentArea.colW1
                      : root.layoutMode === 1 ? contentArea.colW2
                      : contentArea.colW4
                height: col1Content.implicitHeight

                ColumnLayout {
                    id: col1Content
                    anchors.left: parent.left; anchors.right: parent.right; anchors.top: parent.top
                    spacing: AppTheme.spaceXXS

                    Text {
                        text: root.isYearView ? "ВСЕГО ЗА ГОД" : "НА НАЧАЛО МЕСЯЦА"
                        color: AppTheme.textTertiary
                        font.pixelSize: AppTheme.sizeSmall; font.weight: AppTheme.weightBold
                        font.letterSpacing: 1
                        Layout.bottomMargin: AppTheme.spaceXS
                    }
                    SummaryRow {
                        labelText: "ДВО (ночные):"
                        valText: root.isYearView ? (backend.yearSummary.acc_hours || "—")
                                                 : (backend.monthSummary.start_hours || "—")
                    }
                    SummaryRow {
                        labelText: "ДДО (дни):"
                        valText: root.isYearView ? (backend.yearSummary.acc_days || "—")
                                                 : (backend.monthSummary.start_days || "—")
                    }
                    SummaryRow {
                        labelText: "Сверх нормы:"
                        valText: root.isYearView ? (backend.yearSummary.acc_overtime || "—")
                                                 : (backend.monthSummary.start_overtime || "—")
                    }
                }
            }

            // ------------------------------------------
            // КОЛОНКА 2
            // ------------------------------------------
            ColWrapper {
                id: col2Wrapper
                x: root.layoutMode === 0 ? 0
                 : root.layoutMode === 1 ? contentArea.colW2 + AppTheme.spaceXL
                 : contentArea.colW4 + AppTheme.spaceXL
                y: root.layoutMode === 0 ? col1Content.implicitHeight + AppTheme.spaceL : 0
                width:  root.layoutMode === 0 ? contentArea.colW1
                      : root.layoutMode === 1 ? contentArea.colW2
                      : contentArea.colW4
                height: col2Content.implicitHeight

                ColumnLayout {
                    id: col2Content
                    anchors.left: parent.left; anchors.right: parent.right; anchors.top: parent.top
                    spacing: AppTheme.spaceXXS

                    RowLayout {
                        spacing: AppTheme.spaceXS
                        Layout.bottomMargin: AppTheme.spaceXS
                        Layout.fillWidth: true

                        Text {
                            text: root.isYearView ? "ВСЕГО ЗА ГОД" : "ВСЕГО ЗА МЕСЯЦ"
                            color: AppTheme.textTertiary
                            font.pixelSize: AppTheme.sizeSmall; font.weight: AppTheme.weightBold
                            font.letterSpacing: 1
                            Layout.alignment: Qt.AlignVCenter
                        }
                        Item {
                            width: 14; height: 14
                            Layout.alignment: Qt.AlignVCenter
                            property bool isActive: !root.isYearView
                                                    && (backend.monthSummary.is_shift === true)
                            opacity: isActive ? 1.0 : 0.0
                            Behavior on opacity { NumberAnimation { duration: 200 } }
                            IconImage {
                                anchors.centerIn: parent
                                source: "../icons/help.svg"
                                width: 12
                                height: 12
                                color: helpMouseArea.containsMouse ? AppTheme.accentBrand : AppTheme.textTertiary
                            }
                            MouseArea {
                                id: helpMouseArea; anchors.fill: parent; hoverEnabled: true
                                enabled: parent.isActive
                            }
                            AppToolTip {
                                anchors.horizontalCenter: parent.horizontalCenter
                                anchors.bottom: parent.top; anchors.bottomMargin: AppTheme.spaceS
                                isVisible: helpMouseArea.containsMouse && helpMouseArea.enabled
                                text: "Норма за месяц: " + (backend.monthSummary.norm_minutes || "0")
                            }
                        }
                        Item { Layout.fillWidth: true }
                    }

                    SummaryRow {
                        labelText: "ДВО (ночные):"
                        valText: root.isYearView ? (backend.yearSummary.comp_hours || "—")
                                                 : (backend.monthSummary.acc_hours || "—")
                        valColor: root.isYearView ? AppTheme.accentTeal : AppTheme.textPrimary
                    }
                    SummaryRow {
                        labelText: "ДДО (дни):"
                        valText: root.isYearView ? (backend.yearSummary.comp_days || "—")
                                                 : (backend.monthSummary.acc_days || "—")
                        valColor: root.isYearView ? AppTheme.accentTeal : AppTheme.textPrimary
                    }
                    SummaryRow {
                        labelText: "Сверх нормы:"
                        valText: root.isYearView ? (backend.yearSummary.comp_overtime || "—")
                                                 : (backend.monthSummary.acc_overtime || "—")
                        valColor: root.isYearView ? AppTheme.accentTeal : AppTheme.textPrimary
                    }
                    RowLayout {
                        Layout.fillWidth: true; implicitHeight: 18; spacing: AppTheme.spaceXS
                        opacity: (root.isYearView || (backend.monthSummary.is_shift === true)) ? 1.0 : 0.0
                        Behavior on opacity { NumberAnimation { duration: 200 } }
                        Text {
                            Layout.fillWidth: true; Layout.alignment: Qt.AlignVCenter
                            text: root.isYearView ? "Деньгами:" : "График (в ночь):"
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                            elide: Text.ElideRight
                        }
                        OdometerValue {
                            Layout.preferredWidth: 90; Layout.minimumWidth: 60; implicitHeight: 18
                            text: root.isYearView ? (backend.yearSummary.comp_money || "—")
                                                  : (backend.monthSummary.shift_night || "—")
                            textColor: root.isYearView ? AppTheme.textPrimary : AppTheme.accentBrand
                        }
                    }
                    SummaryRow {
                        opacity: (!root.isYearView && (backend.monthSummary.is_shift === true)) ? 1.0 : 0.0
                        Behavior on opacity { NumberAnimation { duration: 200 } }
                        labelText: "График (праздничные):"
                        valText: backend.monthSummary.shift_holiday || "—"
                        valColor: AppTheme.accentPurple
                    }
                }
            }

            // ------------------------------------------
            // КОЛОНКА 3
            // ------------------------------------------
            ColWrapper {
                id: col3Wrapper
                x: root.layoutMode === 0 ? 0
                 : root.layoutMode === 1 ? 0
                 : (contentArea.colW4 + AppTheme.spaceXL) * 2
                y: root.layoutMode === 0
                       ? col1Content.implicitHeight + AppTheme.spaceL
                         + col2Content.implicitHeight + AppTheme.spaceL
                 : root.layoutMode === 1 ? contentArea.row2Y
                 : 0
                width:  root.layoutMode === 0 ? contentArea.colW1
                      : root.layoutMode === 1 ? contentArea.colW2
                      : contentArea.colW4
                height: col3Content.implicitHeight

                ColumnLayout {
                    id: col3Content
                    anchors.left: parent.left; anchors.right: parent.right; anchors.top: parent.top
                    spacing: AppTheme.spaceXXS

                    RowLayout {
                        Layout.fillWidth: true; Layout.bottomMargin: AppTheme.spaceXS
                        implicitHeight: 20

                        Text {
                            Layout.alignment: Qt.AlignVCenter
                            text: root.isYearView ? "СТАТУСЫ ДНЕЙ" : "КОМПЕНСИРОВАНО"
                            color: AppTheme.textTertiary
                            font.pixelSize: AppTheme.sizeSmall; font.weight: AppTheme.weightBold
                            font.letterSpacing: 1
                        }
                        Rectangle {
                            visible: !root.isYearView
                            width: moneyBtnText.implicitWidth + 16
                            height: 22
                            radius: AppTheme.radiusSmall
                            Layout.alignment: Qt.AlignVCenter
                            
                            color: moneyBtnArea.pressed   ? AppTheme.statePress
                                 : moneyBtnArea.containsMouse ? AppTheme.bgBrandSoft
                                 : AppTheme.bgBase
                            border.color: moneyBtnArea.containsMouse ? AppTheme.accentBrand : AppTheme.borderInput
                            border.width: 1
                            
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                            Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }
                            
                            Text {
                                id: moneyBtnText
                                anchors.centerIn: parent
                                text: "₽ Деньги"
                                color: moneyBtnArea.containsMouse ? AppTheme.accentBrand : AppTheme.textSecondary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                font.weight: AppTheme.weightBold
                            }
                            MouseArea {
                                id: moneyBtnArea
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onClicked: { backend.loadMoneyComps(); moneyInspector.show() }
                            }
                            AppToolTip {
                                anchors.horizontalCenter: parent.horizontalCenter
                                anchors.bottom: parent.top
                                anchors.bottomMargin: AppTheme.spaceXXS
                                isVisible: moneyBtnArea.containsMouse
                                text: "Посмотреть денежные компенсации"
                            }
                        }
                        Item { Layout.fillWidth: true }
                    }

                    SummaryRow {
                        labelText: root.isYearView ? "Больничный (Б):" : "ДВО (ночные):"
                        valText: root.isYearView ? (backend.yearSummary.b_days || "—")
                                                 : (backend.monthSummary.comp_hours || "—")
                        valColor: root.isYearView ? AppTheme.accentDanger : AppTheme.accentTeal
                    }
                    SummaryRow {
                        labelText: root.isYearView ? "Отпуск (О):" : "ДДО (дни):"
                        valText: root.isYearView ? (backend.yearSummary.o_days || "—")
                                                 : (backend.monthSummary.comp_days || "—")
                        valColor: root.isYearView ? AppTheme.accentWarning : AppTheme.accentTeal
                    }
                    SummaryRow {
                        labelText: root.isYearView ? "Командировка (К):" : "Сверх нормы:"
                        valText: root.isYearView ? (backend.yearSummary.k_days || "—")
                                                 : (backend.monthSummary.comp_overtime || "—")
                        valColor: root.isYearView ? AppTheme.accentPurple : AppTheme.accentTeal
                    }
                    RowLayout {
                        Layout.fillWidth: true; implicitHeight: 18; spacing: AppTheme.spaceXS
                        opacity: !root.isYearView ? 1.0 : 0.0
                        Behavior on opacity { NumberAnimation { duration: 200 } }
                        Text {
                            Layout.fillWidth: true; Layout.alignment: Qt.AlignVCenter
                            text: "Деньгами:"; color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeSmall
                            elide: Text.ElideRight
                        }
                        OdometerValue {
                            Layout.preferredWidth: 90; Layout.minimumWidth: 60; implicitHeight: 18
                            text: backend.monthSummary.comp_money || "—"
                            textColor: AppTheme.textPrimary
                        }
                    }
                }
            }

            // ------------------------------------------
            // КОЛОНКА 4
            // ------------------------------------------
            ColWrapper {
                id: col4Wrapper
                x: root.layoutMode === 0 ? 0
                 : root.layoutMode === 1 ? contentArea.colW2 + AppTheme.spaceXL
                 : (contentArea.colW4 + AppTheme.spaceXL) * 3
                y: root.layoutMode === 0
                       ? col1Content.implicitHeight + AppTheme.spaceL
                         + col2Content.implicitHeight + AppTheme.spaceL
                         + col3Content.implicitHeight + AppTheme.spaceL
                 : root.layoutMode === 1 ? contentArea.row2Y
                 : 0
                width:  root.layoutMode === 0 ? contentArea.colW1
                      : root.layoutMode === 1 ? contentArea.colW2
                      : contentArea.colW4
                height: col4Content.implicitHeight

                ColumnLayout {
                    id: col4Content
                    anchors.left: parent.left; anchors.right: parent.right; anchors.top: parent.top
                    spacing: AppTheme.spaceXXS

                    Text {
                        text: root.isYearView ? "ВСЕГО НА КОНЕЦ ГОДА" : "НА КОНЕЦ МЕСЯЦА"
                        color: AppTheme.textTertiary
                        font.pixelSize: AppTheme.sizeSmall; font.weight: AppTheme.weightBold
                        font.letterSpacing: 1
                        Layout.bottomMargin: AppTheme.spaceXS
                    }
                    SummaryRow {
                        labelText: "ДВО (ночные):"
                        valText: root.isYearView ? (backend.yearSummary.end_hours || "—")
                                                 : (backend.monthSummary.end_hours || "—")
                        valColor: (root.isYearView ? backend.yearSummary.is_hours_negative
                                                   : backend.monthSummary.is_hours_negative)
                                  ? AppTheme.accentDanger : AppTheme.textPrimary
                    }
                    SummaryRow {
                        labelText: "ДДО (дни):"
                        valText: root.isYearView ? (backend.yearSummary.end_days || "—")
                                                 : (backend.monthSummary.end_days || "—")
                        valColor: (root.isYearView ? backend.yearSummary.is_days_negative
                                                   : backend.monthSummary.is_days_negative)
                                  ? AppTheme.accentDanger : AppTheme.textPrimary
                    }
                    SummaryRow {
                        labelText: "Сверх нормы:"
                        valText: root.isYearView ? (backend.yearSummary.end_overtime || "—")
                                                 : (backend.monthSummary.end_overtime || "—")
                        valColor: (root.isYearView ? backend.yearSummary.is_overtime_negative
                                                   : backend.monthSummary.is_overtime_negative)
                                  ? AppTheme.accentDanger : AppTheme.textPrimary
                    }
                }
            }
        }

        // ════════════════════════════════════════════════════════════
        // МЕСЯЧНЫЙ ВИД (новая «карточная» раскладка)
        // Три показателя (ночные часы / дни / сверх нормы) — как три
        // карточки с иконкой и цветом; внутри — ключевое число
        // «остаток на конец месяца» и мелкий поток начисления.
        // ════════════════════════════════════════════════════════════
        Item {
            id: monthArea
            visible: !root.isYearView
            anchors.top: parent.top
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.margins: AppTheme.spaceM

            // Маленькая «ячейка» потока внутри карточки
            component MiniStat: Column {
                id: ms
                property string valText: "—"
                property string labelText: ""
                property color valColor: AppTheme.textPrimary
                spacing: 1

                Text {
                    width: ms.width
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
                    text: ms.labelText
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeMicro
                    elide: Text.ElideRight
                }
            }

            // Карточка одного показателя
            component CurrencyCard: Rectangle {
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
                Layout.preferredHeight: 150
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

                    RowLayout {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 28
                        Layout.topMargin: AppTheme.spaceXXS
                        spacing: AppTheme.spaceS

                        MiniStat { Layout.fillWidth: true; Layout.fillHeight: true; valText: card.startText; labelText: "на начало" }
                        MiniStat { Layout.fillWidth: true; Layout.fillHeight: true; valText: card.accText;  labelText: "+ начислено"; valColor: AppTheme.accentSuccess }
                        MiniStat { Layout.fillWidth: true; Layout.fillHeight: true; valText: card.compText; labelText: "− компенсир."; valColor: AppTheme.accentDanger }
                    }
                }
            }

            ColumnLayout {
                id: monthBody
                anchors.top: parent.top
                anchors.left: parent.left
                anchors.right: parent.right
                spacing: AppTheme.spaceM

                // Заголовок панели
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
                        Layout.fillWidth: true
                        spacing: 0
                        Text {
                            text: "Балансы"
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBodyLarge
                            font.weight: AppTheme.weightBold
                        }
                        Text {
                            text: backend.currentPeriodText
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                    }

                    // Кнопка «Деньги» — открывает просмотр денежных компенсаций
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
                }

                // Три карточки показателей (часы / дни / сверх нормы)
                RowLayout {
                    id: cardsRow
                    Layout.fillWidth: true
                    spacing: AppTheme.spaceM

                    CurrencyCard {
                        title: "ДВО (ночные)"
                        icon: "clock.svg"
                        accent: AppTheme.accentBrand
                        endText: backend.monthSummary.end_hours || "—"
                        endNeg: backend.monthSummary.is_hours_negative === true
                        startText: backend.monthSummary.start_hours || "—"
                        accText: backend.monthSummary.acc_hours || "—"
                        compText: backend.monthSummary.comp_hours || "—"
                    }
                    CurrencyCard {
                        title: "ДДО (дни)"
                        icon: "calendar.svg"
                        accent: AppTheme.accentSuccess
                        endText: backend.monthSummary.end_days || "—"
                        endNeg: backend.monthSummary.is_days_negative === true
                        startText: backend.monthSummary.start_days || "—"
                        accText: backend.monthSummary.acc_days || "—"
                        compText: backend.monthSummary.comp_days || "—"
                    }
                    CurrencyCard {
                        title: "Сверх нормы"
                        icon: "sparkle.svg"
                        accent: AppTheme.accentWarning
                        endText: backend.monthSummary.end_overtime || "—"
                        endNeg: backend.monthSummary.is_overtime_negative === true
                        startText: backend.monthSummary.start_overtime || "—"
                        accText: backend.monthSummary.acc_overtime || "—"
                        compText: backend.monthSummary.comp_overtime || "—"
                    }
                }

                // Строка-подсказка для сменного графика
                RowLayout {
                    Layout.fillWidth: true
                    visible: backend.monthSummary.is_shift === true
                    spacing: AppTheme.spaceS

                    IconImage {
                        source: "../icons/help.svg"
                        width: 13; height: 13
                        color: AppTheme.textTertiary
                        Layout.alignment: Qt.AlignVCenter
                    }
                    Text {
                        Layout.fillWidth: true
                        Layout.alignment: Qt.AlignVCenter
                        text: "Сменный график: в ночь " + (backend.monthSummary.shift_night || "—")
                              + "  ·  праздничные " + (backend.monthSummary.shift_holiday || "—")
                              + "  ·  норма за месяц " + (backend.monthSummary.norm_minutes || "0")
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        wrapMode: Text.WordWrap
                    }
                }
            }
        }
    }
}
