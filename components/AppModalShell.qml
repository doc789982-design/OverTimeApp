import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

// ============================================================
// ОБОЛОЧКА МОДАЛЬНОГО ОКНА — общая база всех модалок программы
// (AppDialog, AppLargeModal, окна календаря и рекомендаций).
// Здесь и только здесь живут: затемнение фона с анимацией,
// вход/выход (одни и те же тайминги у всех окон), фон с тенью,
// страховка закрытия и слой. Разным окнам — разный контент,
// но не разная механика появления.
// ============================================================
Popup {
    id: root

    // Клик мимо окна: формы НЕ закрывают (случайный щелчок не должен
    // сжигать введённое — стандарт Apple/Google), читалки — закрывают.
    property bool closeOnOutside: false

    // Режим «морфинга»: окно открывается сразу в заданном месте
    // без анимации масштаба (продолжает рост ячейки дня).
    property bool morphOpen: false

    // Фону окна нужен clip (контент не должен вылезать за скруглённые
    // углы) — включается окнами с большим контентом (AppLargeModal).
    property bool clipBackground: false

    z: AppTheme.zModal
    modal: true
    dim: true
    focus: true
    closePolicy: root.closeOnOutside
        ? (Popup.CloseOnEscape | Popup.CloseOnPressOutside)
        : Popup.CloseOnEscape

    // Попап — Item в overlay главного окна, а не нативное ОС-окно:
    // нативное окно с модальностью и анимациями закрывается нестабильно
    // (может оставить невидимую зону, блокирующую клики)
    popupType: Popup.Item

    // Страховка: если выходная анимация по какой-то причине не завершилась,
    // принудительно докрываем попап — иначе его невидимый прямоугольник
    // блокирует клики в зоне, где окно было
    Timer {
        id: closeSafety
        interval: 600
        onTriggered: if (!root.opened) root.visible = false
    }
    onAboutToHide: closeSafety.restart()
    // Окно открыли снова — страховка закрытия отменяется. Без этого она
    // могла выстрелить посреди входного перехода нового открытия
    // (opened становится true только ПОСЛЕ перехода) и убить окно:
    // быстро закрыли и открыли — и диалог завис видимым, но не открытым
    onAboutToShow: closeSafety.stop()

    // Затемнение фона — единое у всех модальных окон
    Overlay.modal: Rectangle {
        color: AppTheme.bgOverlay
        opacity: root.opened ? 1.0 : 0.0
        Behavior on opacity {
            NumberAnimation { duration: AppTheme.durStandard; easing.type: root.opened ? AppTheme.easeEnter : AppTheme.easeExit }
        }
    }

    // Вход: окно проявляется быстро и «вырастает» из 92%.
    // При морфинге масштаб уже 1.0 (from == to) — анимации роста нет,
    // окно просто продолжает рост ячейки как единое целое.
    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: root.morphOpen ? 1.0 : 0.92; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }
    }
    // Выход: быстрое торможение (вход — разгон, выход — торможение)
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        color: AppTheme.bgModal
        radius: AppTheme.radiusModal
        border.color: AppTheme.borderDivider
        border.width: 1
        clip: root.clipBackground
        // Тень-картинка вместо вычисляемой (легко для видеокарты)
        AppShadow { level: 4 }
    }
}
