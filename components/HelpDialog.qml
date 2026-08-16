import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

AppSidePanel {
    id: root
    width: 550 // Сделали чуть шире для комфортного чтения
    title: "О программе и справка"

    Column {
        width: parent.width
        spacing: AppTheme.spaceL

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
        // БЛОК 2: КРАТКАЯ СПРАВКА (Старый текст)
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
    }
}