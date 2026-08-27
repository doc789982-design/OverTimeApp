import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

AppSidePanel {
    id: root
    width: 550 // Сделали чуть шире для комфортного чтения
    title: "О программе и справка"

    readonly property string githubUrl: "https://github.com/doc789982-design/OverTimeApp/releases/latest"

    // Человекочитаемое описание действия горячей клавиши.
    // Если пользователь дал имени название — показываем его, иначе собираем сами.
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

    Column {
        width: parent.width
        spacing: AppTheme.spaceL

        // Версия программы (единственный источник — AppTheme.appVersion)
        Text {
            width: parent.width
            text: "OVERTIMETAB · " + AppTheme.appVersionFull
            color: AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightMedium
            horizontalAlignment: Text.AlignHCenter
        }

        // ==========================================
        // СКАЧАТЬ АКТУАЛЬНУЮ ВЕРСИЮ (GitHub + QR)
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
                        smooth: false
                        mipmap: false
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
                        text: "github.com/doc789982-design/OverTimeApp"
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

        // ==========================================
        // БЛОК 1: ЮРИДИЧЕСКАЯ ИНФОРМАЦИЯ
        // ==========================================
        Rectangle {
            width: parent.width
            height: licenseColumn.implicitHeight + (AppTheme.spaceM * 2)
            color: AppTheme.bgDangerSoft // Легкий красный оттенок для привлечения внимания
            radius: AppTheme.radiusMedium
            border.color: AppTheme.accentDanger
            border.width: 1

            Column {
                id: licenseColumn
                anchors.fill: parent
                anchors.margins: AppTheme.spaceM
                spacing: AppTheme.spaceS

                Text { 
                    text: "УВЕДОМЛЕНИЕ ОБ АВТОРСКИХ ПРАВАХ И УСЛОВИЯХ ИСПОЛЬЗОВАНИЯ"
                    color: AppTheme.accentDanger
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBodyLarge
                    font.weight: AppTheme.weightBold 
                    width: parent.width
                    wrapMode: Text.WordWrap
                }

                Text { 
                    text: "<b>1. СТАТУС ПРОГРАММЫ</b><br>Настоящая сборка программного обеспечения является закрытой тестовой версией. Программа предоставляется исключительно в целях ознакомления, тестирования работоспособности и выявления возможных технических ошибок.<br><br>" +
                          "<b>2. ИНТЕЛЛЕКТУАЛЬНАЯ СОБСТВЕННОСТЬ</b><br>Все исключительные права на данное программное обеспечение, включая исходный код, алгоритмы расчетов, дизайн интерфейса и архитектуру баз данных, являются интеллектуальной собственностью и принадлежат исключительно ее автору-разработчику:<br>" +
                          "<font color='" + AppTheme.accentBrand + "'><b>Специалисту направления профессиональной подготовки отделения по работе с личным составом ОМВД России по г. Мичуринску лейтенанту полиции Григорьеву Максиму Викторовичу</b></font> (далее — Правообладатель).<br><br>" +
                          "<b>3. ОГРАНИЧЕНИЯ ИСПОЛЬЗОВАНИЯ</b><br>Настоящая тестовая версия не предназначена для коммерческого использования или свободного распространения. Запрещается копировать, декомпилировать или передавать программу третьим лицам без согласия Правообладателя.<br><br>" +
                          "<b>4. ОТКАЗ ОТ ОТВЕТСТВЕННОСТИ</b><br>Программа предоставляется на условиях «как есть» (as is). Правообладатель не несет ответственности за потерю данных, возникшую в процессе тестирования."
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    textFormat: Text.RichText
                    width: parent.width
                    wrapMode: Text.WordWrap
                    lineHeight: 1.2
                }
            }
        }

        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

        // ==========================================
        // БЛОК 2: КРАТКАЯ СПРАВКА
        // ==========================================
        Text { 
            text: "Краткое руководство"
            color: AppTheme.textPrimary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeH2
            font.weight: AppTheme.weightBold 
        }

        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Дежурство"; color: AppTheme.accentBrand; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold }
            Text { 
                text: "Двойной клик по дню в календаре — быстрое добавление дежурства. Правый клик по дню открывает меню: статусы, компенсации, «Открыть день» и тип дня (рабочий / выходной / праздничный)."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap 
            }
        }

        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Горячие клавиши"; color: AppTheme.accentBrand; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold }
            Text { 
                text: "Наведите курсор мыши на любой день в календаре (не кликая) и нажмите заданную в настройках комбинацию клавиш."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap 
            }
        }

        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Перевод сотрудников"; color: AppTheme.accentPurple; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold }
            Text { 
                text: "Перетащите карточку на нужную группу слева. Если зажать клавишу SHIFT при перетаскивании, откроется окно 'Официального перевода'."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap 
            }
        }

        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Статусы дней (Б, О, К)"; color: AppTheme.accentWarning; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold }
            Text { 
                text: "Нажмите правой кнопкой мыши по ячейке в календаре для выбора статуса."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap 
            }
        }

        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Отмена действий"; color: AppTheme.accentTeal; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold }
            Text { 
                text: "Ctrl+Z — отменить последнее действие, Ctrl+Y — вернуть отменённое. Стек истории — 20 шагов."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap 
            }
        }

        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Обновление"; color: AppTheme.accentBrand; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBodyLarge; font.weight: AppTheme.weightBold }
            Text {
                text: "Новую версию не нужно ставить поверх старой вручную. Положите zip рядом с OVERTIMETAB.exe (или на флешку) — через несколько секунд внизу слева появится кнопка «Обновить». Базы и горячие клавиши сохранятся. Файл можно указать и в Настройках → Обновление. После установки программа покажет, что изменилось."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap
            }
        }

        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

        // ==========================================
        // БЛОК 3: ВАШИ ГОРЯЧИЕ КЛАВИШИ (ЖИВОЙ СПИСОК)
        // ==========================================
        Column {
            width: parent.width; spacing: AppTheme.spaceXS
            Text { text: "Ваши горячие клавиши"; color: AppTheme.textPrimary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeH3; font.weight: AppTheme.weightBold }
            Text { 
                text: "Список берётся из ваших настроек. Изменить: «Настройки» → «Горячие клавиши»."
                color: AppTheme.textSecondary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; width: parent.width; wrapMode: Text.WordWrap 
            }
        }

        Column {
            width: parent.width
            spacing: AppTheme.spaceS
            visible: backend.hotkeysList.length > 0

            Repeater {
                model: backend.hotkeysList

                RowLayout {
                    width: parent.width
                    spacing: AppTheme.spaceM

                    // Бейдж с клавишей
                    Rectangle {
                        Layout.preferredWidth: Math.max(44, hkKeyText.implicitWidth + 16)
                        Layout.preferredHeight: 26
                        radius: AppTheme.radiusSmall
                        color: AppTheme.bgBase
                        border.color: AppTheme.borderInput
                        border.width: 1

                        Text {
                            id: hkKeyText
                            anchors.centerIn: parent
                            text: modelData.key
                            color: AppTheme.accentBrand
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                        }
                    }

                    Text {
                        Layout.fillWidth: true
                        text: root.describeHotkey(modelData)
                        color: AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        wrapMode: Text.WordWrap
                    }
                }
            }
        }

        // Подсказка, если клавиш ещё нет
        Rectangle {
            width: parent.width
            height: emptyHkText.implicitHeight + AppTheme.spaceM * 2
            visible: backend.hotkeysList.length === 0
            radius: AppTheme.radiusMedium
            color: AppTheme.bgSurface
            border.color: AppTheme.borderDivider
            border.width: 1

            Text {
                id: emptyHkText
                anchors.centerIn: parent
                width: parent.width - AppTheme.spaceL * 2
                text: "Горячие клавиши пока не настроены. Создайте первую в разделе «Настройки» → «Горячие клавиши» — и заполняйте табель одним нажатием."
                color: AppTheme.textTertiary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                horizontalAlignment: Text.AlignHCenter
                wrapMode: Text.WordWrap
            }
        }
    }
}
