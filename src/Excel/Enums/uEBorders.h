#ifndef UEBORDERS_H
#define UEBORDERS_H

//---------------------------------------------------------------------------

/// @brief Индекс для обращения к границам диапазона
enum class XlBordersIndex : char {
    xlDiagonalDown = 5,         /// Диагональ из верхнего левого
    xlDiagonalUp = 6,           /// Диагональ из нижнего левого 
    xlEdgeBottom = 9,           /// Нижняя
    xlEdgeLeft = 7,             /// Левая
    xlEdgeRight = 10,           /// Правая
    xlEdgeTop = 8,              /// Верхняя
    xlInsideHorizontal = 12,    /// Внутренние горизонтальные
    xlInsideVertical = 11       /// Внутренние вертикальные
};

/// @brief Толщина границ(ы)
enum class XlBorderWeight : short int {
    xlHairline = 1,     /// Очень тонкая (точечная)
    xlMedium = -4138,   /// Жирная
    xlThick = 4,        /// Двойная
    xlThin = 2          /// Обычная, тонкая
};

/// @brief Тип линии
enum class XlLineStyle : short int {
    xlContinuous = 1,	        /// Обычная, непрерывная
    xlDash = -4115,	            /// Длинная пунктирная
    xlDashDot = 4,	            /// Пунктирная с длинными и короткими пунктирами
    xlDashDotDot = 5,	        /// Пунктирная с двумя короткими
    xlDot = -4118,	            /// Короткий пунктир 
    xlDouble = -4119,	        /// Двойная
    xlLineStyleNone = -4142,	/// Без линии
    xlSlantDashDot = 13	        /// Наклонные черточки
};

//---------------------------------------------------------------------------
#endif

