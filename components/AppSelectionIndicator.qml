import QtQuick

Rectangle {
    id: root

    // ==========================================
    // ВХОДНЫЕ СИГНАЛЫ
    // ==========================================
    property bool isHovered: false
    property bool isSelected: false
    property bool isDragged: false
    
    // Базовая высота родительского элемента, от которой мы считаем проценты
    property real targetHeight: 48 

    // ==========================================
    // ДИЗАЙН-ТОКЕНЫ
    // ==========================================
    width: 4 // Сделаем 4px, чтобы он смотрелся уверенно
    radius: 2
    color: AppTheme.accentBrand

    // ==========================================
    // ФИЗИКА ВЫСОТЫ
    // ==========================================
    height: isDragged ? targetHeight * 0.8 :
           (isSelected ? targetHeight * 0.6 : 
           (isHovered ? targetHeight * 0.3 : 0))
    
    Behavior on height { 
        NumberAnimation { 
            duration: AppTheme.durNormal 
            easing.type: AppTheme.easeEnter 
        } 
    }
}