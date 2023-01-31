/** @file 
 * @brief Перечисления сводных таблиц
 * @ingroup Enums
 * 
 * Тут находятся перечисления, необходимые для настройки сводных таблиц
 */
#ifndef uTExcelEnumsPivotH
#define uTExcelEnumsPivotH

//---------------------------------------------------------------------------
/// Источник данных для сводных таблиц
enum class XlPivotTableSourceType : short int {
	xlConsolidation = 3,	/// Несколько диапазонов консолидации
	xlDatabase = 1,			/// Microsoft Excel список или база данных
	xlExternal = 2,			/// Данные из другого приложения
	xlPivotTable = -4148	/// Тот же источник, что и другой отчет PivotTable
};

/// Функции для сводных таблиц
enum class XlConsolidationFunction : short int { 
	xlAverage = -4106,		/// Среднее значение
	xlCount = -4112,		/// Количество
	xlCountNums = -4113,	/// Количество только численных значений
	//xlDistinctCount = 11,	// Количество уникальных значений - Не доступна в старых версиях
	xlMax = -4136,			/// Максимум
	xlMin = -4139,			/// Минимум
	xlProduct = -4149,		/// Произверение
	xlStDev = -4155,		/// Стандартное отклонение, основанное на конкретном значении (?)
	xlStDevP = -4156,		/// Стандартное отклонение, основанное на всех значениях
	xlSum = -4157,			/// Сумма
	xlUnknown = 1000,		/// Функция не указана
	xlVar = -4164,			/// Вариант, основанный на конкретном значении
	xlVarP = -4165			/// Вариант, основанный на всех значениях
};

/// Ориентация полей в сводных таблицах
enum class XlPivotFieldOrientation : unsigned short int {
	xlColumnField = 2,	/// Колонка
	xlDataField = 4,	/// Данные
	xlHidden = 0,		/// Скрыто
	xlPageField = 3,	/// Страничка
	xlRowField = 1		/// Строка
};

//---------------------------------------------------------------------------
#endif  // uTExcelEnumsPivotH

