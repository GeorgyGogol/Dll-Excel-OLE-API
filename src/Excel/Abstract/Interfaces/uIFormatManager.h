#ifndef UIFORMATMANAGER_H
#define UIFORMATMANAGER_H

namespace exl {
/** @addtogroup Interfaces
 * @{
 * 
 * @brief Интерфейс для настройки формата ячеек
 * 
 * Отвечает за добавление методов взаимодействия с форматированием ячейки
 * или ячеек.
 * 
 */
template<class T>
class IFormatManager {
public:
	/// @brief Горизонтальное выравнивание текста в ячейке(ах)
	/// @param align Выравнивание
	virtual T* SetHorizontalAlign(ExcelTextAlign align) = 0;

	/// @brief Горизонтальное выравнивание текста в ячейке(ах)
	/// @param align Выравнивание
	virtual T* SetVerticalAlign(ExcelTextAlign align) = 0;
		
	//virtual T* SetBorders() = 0;
		

	//virtual T* SetWidth() = 0;
	//virtual T* SetHeight() = 0;
	//virtual T* AutoSize() = 0;
		
	//virtual T* SetFormat() = 0;
};

}
/// @}
#endif

