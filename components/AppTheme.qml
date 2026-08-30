pragma Singleton
import QtQuick

QtObject {
    id: theme

    // ==========================================
    // ВЕРСИЯ ПРОГРАММЫ — ЕДИНСТВЕННОЕ МЕСТО
    // Меняем здесь: заголовок окна, шапка и логотип
    // подхватят автоматически.
    // ==========================================
    readonly property string appVersion: "2.0.0-ALPHA.20"
    readonly property int appBuild: 123
    readonly property string appVersionFull: appVersion + " · сборка " + appBuild

    // ==========================================
    // ЗАГРУЗКА ШРИФТОВ (ГАРАНТИЯ ПРАВИЛЬНЫХ ИМЕН)
    // ==========================================
    // Мы сохраняем загрузчики в свойства (с нижним подчеркиванием, чтобы они не мешались),
    // чтобы Qt добавил их в базу и прочитал реальные имена.
    property FontLoader _reg: FontLoader { source: "../fonts/Roboto-Regular.ttf" }
    property FontLoader _med: FontLoader { source: "../fonts/Roboto-Medium.ttf" }
    property FontLoader _bld: FontLoader { source: "../fonts/Roboto-Bold.ttf" }
    
    property FontLoader _condReg: FontLoader { source: "../fonts/RobotoCondensed-Regular.ttf" }
    property FontLoader _condBld: FontLoader { source: "../fonts/RobotoCondensed-Bold.ttf" }

    // Подхватываем тему из Питона
    property bool isDark: backend.isDarkTheme

    // ==========================================
    // 1. ПОВЕРХНОСТИ (Surfaces)
    // ==========================================
    // Светлая — Canvas (InstUI): белая страница, фарфор у панелей, карточки белые.
    // Тёмная — Material: #121212 и ступени elevation, календарь чуть светлее рельс.
    // Тёмная тема — Material dark surfaces (0dp→12dp) с правильным порядком
    // elevation: фон самый тёмный, поднимающиеся поверхности светлеют.
    property color bgBase:     isDark ? "#121212" : "#FFFFFF"
    property color bgPanel:    isDark ? "#121212" : "#F5F5F5"
    property color bgSurface:  isDark ? "#1E1E1E" : "#F8F9FA"   // 1dp
    property color bgCell:     isDark ? "#242424" : "#F4F5F7"   // 3dp
    property color bgElevated: isDark ? "#2C2C2C" : "#FFFFFF"   // 6dp
    property color bgModal:    isDark ? "#333333" : "#FFFFFF"   // 12dp
    property color bgInput:    "transparent"

    property color bgDisabled:       isDark ? "#212529" : "#E9ECEF" 
    property color bgSkeletonBase:   isDark ? "#2B3035" : "#E2E8F0" 
    property color bgSkeletonShine:  isDark ? "#3B4048" : "#F8F9FA" 

    // ==========================================
    // 2. СЕМАНТИЧЕСКИЕ ЦВЕТА (Brand & Status)
    // ==========================================
    // Светлая — палитра Canvas (Workday): blueberry / greenApple / sourLemon /
    // cinnamon / grapeSoda / jewel. Тёмная — приглушённые для тёмного фона.
    property color accentBrand:   isDark ? "#4B9AEE" : "#0875E1"  // blueberry400
    property color accentInfo:    isDark ? "#82B1FF" : "#005CB9"  // blueberry500
    property color accentSuccess: isDark ? "#5CB85C" : "#319C4C"  // greenApple500
    property color accentWarning: isDark ? "#F4AC5B" : "#EBB400"  // sourLemon500
    property color accentDanger:  isDark ? "#E76C78" : "#DE2E21"  // cinnamon500
    property color accentPurple:  isDark ? "#C05C9A" : "#97499E"  // grapeSoda500
    property color accentTeal:    isDark ? "#50A7B5" : "#1A818C"  // jewel500

    // Мягкие фоны (Зашиты прямо в HEX-коды: 33 = 20% opacity, 1E = 12% opacity)
    // Это 100% безопасный метод для движка QML
    property color bgBrandSoft:   isDark ? "#334B9AEE" : "#1E0875E1"

    // Годовая сетка: клетки ~12px. Обычное дежурство — плотный бренд.
    // Сменное — тот же цвет, вполовину прозрачнее.
    property color yearDutyExtra: accentBrand
    property color yearDutyShift: Qt.rgba(accentBrand.r, accentBrand.g, accentBrand.b, 0.5)
    property color yearWeekend:   isDark ? "#66E76C78" : "#4DDE2E21"
    property color bgInfoSoft:    isDark ? "#3382B1FF" : "#1E005CB9"
    property color bgSuccessSoft: isDark ? "#335CB85C" : "#1E319C4C"
    property color bgWarningSoft: isDark ? "#33F4AC5B" : "#1EEBB400"
    property color bgDangerSoft:  isDark ? "#33E76C78" : "#1EDE2E21"
    property color bgPurpleSoft:  isDark ? "#33C05C9A" : "#1E97499E"
    property color bgTealSoft:    isDark ? "#3350A7B5" : "#1E1A818C"

    // ==========================================
    // 3. ТЕКСТ (Typography Colors)
    // ==========================================
    property color textPrimary:   isDark ? "#DEE4EA" : "#2D3B45"
    property color textSecondary: isDark ? "#8B949E" : "#6B7780"
    property color textTertiary:  isDark ? "#6B757D" : "#8B969E"
    property color textDisabled:  isDark ? "#484F58" : "#9AA0A6" 
    
    property color textInverse:   isDark ? "#202124" : "#FFFFFF" 
    property color textOnAccent:  "#FFFFFF" 
    property color textOnSoft:    accentBrand 

    // ==========================================
    // 4. ТИПОГРАФИКА (Strict Scale & Weights)
    // ==========================================
    // МАГИЯ: Берем точные имена прямо из файлов, чтобы Qt не путался!
    property string fontFamily: _reg.name 
    property string fontCondensed: _condReg.name 

    // Градация толщины (Начертания)
    property int weightLight:   300 
    property int weightRegular: 400 
    property int weightMedium:  500 
    property int weightBold:    700 
    property int weightBlack:   900 

    property int sizeH1:        40
    property int sizeH2:        32
    property int sizeH3:        28
    property int sizeH4:        24
    property int sizeH5:        20
    property int sizeBodyLarge: 16
    property int sizeBody:      14 
    property int sizeSmall:     12 
    property int sizeMicro:     10 

    // Межстрочные интервалы и трекинг (как у type-уровней Canvas)
    property real lineHeightTight:  1.2
    property real lineHeightBody:   1.5
    property real lineHeightLoose:  1.6
    property real letterSpacingNone: 0.0
    property real letterSpacingWide: 0.6
    property real letterSpacingCaps: 1.0

    // ==========================================
    // 5. ГРАНИЦЫ И ФОКУС (Borders & Accessibility)
    // ==========================================
    property color borderDivider:  isDark ? "#30363D" : "#E8EAEC"
    property color borderInput:    isDark ? "#484F58" : "#C7CDD1" 
    property color borderDisabled: isDark ? "#353B42" : "#DDE2E5" 
    property color borderError:    accentDanger 
    
    // Фокус сплошным цветом (как у Canvas: 2px + отступ)
    property color borderFocus:    isDark ? "#4B9AEE" : "#0875E1"
    
    property int   focusWidth:     2 
    property int   focusOffset:    2 

    // ==========================================
    // 6. ИНТЕРАКТИВ И ПРОЗРАЧНОСТЬ (States & Opacity)
    // ==========================================
    property color stateHover:    isDark ? Qt.rgba(1, 1, 1, 0.06) : Qt.rgba(0, 0, 0, 0.04)
    property color statePress:    isDark ? Qt.rgba(1, 1, 1, 0.12) : Qt.rgba(0, 0, 0, 0.08)
    property color stateSelected: isDark ? "#264B9AEE" : "#1A0875E1" // 15% и 10% opacity

    property color bgOverlay:     isDark ? Qt.rgba(0, 0, 0, 0.6) : Qt.rgba(32/255, 33/255, 36/255, 0.4) 
    
    property real  alphaDisabled: 0.38 
    property real  scaleActive:   0.98 

    // ==========================================
    // 7. СКРУГЛЕНИЯ (Border Radii)
    // ==========================================
    // По ролям Canvas shape: small 4 (компактные), medium 8 (формы),
    // large 16 (карточки xl), modal 24 (диалоги xxl), pill 999 (full).
    property int radiusSharp:  0    
    property int radiusSmall:  4    
    property int radiusMedium: 8    
    property int radiusLarge:  16   
    property int radiusModal:  24   
    property int radiusPill:   999  

    // ==========================================
    // 8. СЕТКА И ОТСТУПЫ (Spacing Scale)
    // ==========================================
    property int spaceMicro: 2
    property int spaceXXS:   4
    property int spaceXS:    8
    property int spaceS:     12
    property int spaceM:     16 
    property int spaceL:     24
    property int spaceXL:    32
    property int spaceXXL:   48
    property int spaceXXXL:  64

    // Одинаковые по смыслу ряды и шапки — одни числа.
    property int rowHeight:         56
    property int barHeight:         60
    property int fieldIconSize:     36
    property int cardActionReserve: 84
    property int startCardHeight:   72

    // ==========================================
    // 9. ГЛУБИНА И ТЕНИ (Elevation Levels - Material 3)
    // ==========================================
    // Цвет тени: 12% черного днем, 40% черного ночью (чтобы не было эффекта "грязи")
    property color shadowColor: isDark ? Qt.rgba(0, 0, 0, 0.40) : Qt.rgba(0, 0, 0, 0.12)
    
    // Level 1 (Raised) - Карточки, кнопки
    property int shadowL1Y:     1
    property int shadowL1Blur:  3
    
    // Level 2 (Overlay) - Выпадающие меню
    property int shadowL2Y:     3
    property int shadowL2Blur:  6
    
    // Level 3 (Sticky) - Плавающие шапки (или элементы при перетаскивании)
    property int shadowL3Y:     4
    property int shadowL3Blur:  8
    
    // Level 4 (Modal) - Диалоговые окна
    property int shadowL4Y:     8
    property int shadowL4Blur:  12
    
    // Level 5 (Pop-out) - Тултипы, Тоасты (Уведомления)
    // В Material у тултипов очень плотная и направленная вниз тень
    property int shadowL5Y:     12
    property int shadowL5Blur:  16

    // ==========================================
    // 10. Z-СЛОИ (Z-Index Architecture)
    // ==========================================
    property int zBackground: 0      
    property int zContent:    10     
    property int zSticky:     100    
    property int zOverlay:    9000   
    property int zModal:      9010   
    property int zDropdown:   9050   
    property int zTooltip:    9900   
    property int zToast:      9990   
    property int zEffect:     99999  

    // ==========================================
    // 11. ИКОНКИ (Icon Sizes)
    // ==========================================
    property int iconSmall:  16
    property int iconMedium: 16
    property int iconLarge:  20
    property int iconXL:     24
    property int iconXXL:    32

    // ==========================================
    // 12. АНИМАЦИИ И ФИЗИКА (Motion System)
    // ==========================================
    property int durMicro:    100 
    property int durFast:     150 
    property int durNormal:   200  // Чуть быстрее
    
    // Стандарт для окон: 250мс (оптимально для экспоненты, быстро и плавно)
    property int speedStandard: 50
    property int durStandard: 250 
    property int durSlow:     400 

    // Мягкие кривые (как у Canvas/Material) вместо резких экспонент
    property int easeColor:    Easing.Linear     
    property int easeEnter:    Easing.OutCubic  
    property int easeExit:     Easing.InCubic   
    property int easeStandard: Easing.InOutCubic

    property int slideOffset:  20 // Чуть уменьшили разбег, чтобы соответствовало новой скорости

    // ==========================================
    // 13. АДАПТИВНОСТЬ И СЕТКА (Breakpoints)
    // ==========================================
    property int bpCompact:  600   
    property int bpMedium:   1024  
    property int bpExpanded: 1440  
    
    property int maxContentWidth: 1200 
}