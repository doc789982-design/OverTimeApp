import QtQuick
import "."

// ============================================================
// МИНИ-ЯЧЕЙКА ДНЯ для полосы месяца напротив сотрудника.
//
// Общая картина подразделения: когда сотрудник не выбран, карточка
// каждого человека продолжается строкой таких ячеек — те же цвета
// и метки, что у ячейки календаря, только меньше: номер дня, фон
// выходного/праздника, запертые дни приглушены, круглая метка
// статуса (К/Б/О), метка компенсации «В» и цветные полоски
// дежурств внизу (синяя — по графику сменности, серая — суточное).
// ============================================================
Rectangle {
    id: cell

    property var dayData: null

    readonly property bool locked: dayData !== null
                                   && (dayData.is_before_hire === true
                                       || dayData.is_after_end === true)
    readonly property bool isToday: {
        if (dayData === null) return false
        let d = new Date()
        return dayData.date_str === (d.getFullYear() + "-" +
                ("0" + (d.getMonth() + 1)).slice(-2) + "-" +
                ("0" + d.getDate()).slice(-2))
    }

    radius: AppTheme.radiusSmall
    color: dayData !== null && (dayData.is_weekend || dayData.is_holiday)
           ? AppTheme.bgDangerSoft
           : AppTheme.bgCell
    opacity: locked ? 0.3 : 1.0
    border.width: isToday ? 1 : 0
    border.color: AppTheme.accentBrand

    // Номер дня — как в календаре: мелкий, полужирный
    Text {
        anchors.horizontalCenter: parent.horizontalCenter
        anchors.top: parent.top
        anchors.topMargin: 1
        text: cell.dayData !== null ? cell.dayData.day_number : ""
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeMicro
        font.weight: AppTheme.weightBold
        color: cell.isToday ? AppTheme.accentBrand : AppTheme.textSecondary
    }

    // Круглые метки: статус (К/Б/О) и компенсация «В» —
    // те же цвета, что у бейджей календаря
    Row {
        anchors.centerIn: parent
        spacing: 1

        Rectangle {
            visible: cell.dayData !== null && cell.dayData.status !== ""
            width: 12; height: 12; radius: 6
            color: !visible ? "transparent"
                   : (cell.dayData.status === "Б" ? AppTheme.bgDangerSoft
                      : (cell.dayData.status === "О" ? AppTheme.bgWarningSoft
                         : AppTheme.bgPurpleSoft))
            Text {
                anchors.centerIn: parent
                text: cell.dayData !== null ? cell.dayData.status : ""
                font.family: AppTheme.fontFamily
                font.pixelSize: 7
                font.weight: AppTheme.weightBold
                color: cell.dayData === null ? "transparent"
                       : (cell.dayData.status === "Б" ? AppTheme.accentDanger
                          : (cell.dayData.status === "О" ? AppTheme.accentWarning
                             : AppTheme.accentPurple))
            }
        }

        Rectangle {
            visible: cell.dayData !== null && cell.dayData.has_comp === true
            width: 12; height: 12; radius: 6
            color: AppTheme.bgTealSoft
            Text {
                anchors.centerIn: parent
                text: "В"
                font.family: AppTheme.fontFamily
                font.pixelSize: 7
                font.weight: AppTheme.weightBold
                color: AppTheme.accentTeal
            }
        }
    }

    // Дежурства — цветные полоски внизу (пилюльки календаря
    // без текста: в мини-ячейке времени не хватает)
    Column {
        anchors.horizontalCenter: parent.horizontalCenter
        anchors.bottom: parent.bottom
        anchors.bottomMargin: 2
        spacing: 1

        Repeater {
            model: cell.dayData !== null ? Math.min(cell.dayData.duties.length, 3) : 0
            Rectangle {
                width: cell.width - 6
                height: 3
                radius: 1.5
                color: cell.dayData.duties[index].is_shift
                       ? AppTheme.accentBrand : AppTheme.textTertiary
            }
        }
    }
}
