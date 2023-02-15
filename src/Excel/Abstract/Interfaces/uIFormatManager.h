#ifndef UIFORMATMANAGER_H
#define UIFORMATMANAGER_H
/** @file
 * @brief Содержит Интерфейс настройки форматирования
 */

#include "uEnums.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Интерфейс для настройки формата ячеек
 * @ingroup Interfaces
 * 
 * Отвечает за добавление методов взаимодействия с форматированием ячейки
 * или ячеек.
 * 
 * ---------------------------------------------------------------------- **/
template<class T>
class IFormatManager 
{
public:
	/// @brief Получить интерфейс форматирования
	/// @return Интерфейс текущего объекта
	virtual IFormatManager<T>* GetFormatInterface() = 0;

	// @brief Объеденить ячейки
	//virtual T* Merge() = 0;

	/// @brief Горизонтальное выравнивание текста в ячейке(ах)
	/// @param align Выравнивание
	virtual T* SetHorizontalAlign(ExcelTextAlign align) = 0;

	/// @brief Вертикальное выравнивание текста в ячейке(ах)
	/// @param align Выравнивание
	virtual T* SetVerticalAlign(ExcelTextAlign align) = 0;

	/// @brief Автоперенос текста в ячейке
	/// @param state Делать или нет
	virtual T* SetTextWrap(bool state) = 0;
	
	/// @brief Установить один из способов форматирования ячейки
	virtual T* SetFormat(ExcelFormats format) = 0;
	// OlePropertySet("NumberFormat", "@");

	/// @brief Установить свой способ форматирования ячейки
	/// @param format Строка, по которой будет происходить форматирование
	virtual T* SetFormat(const String& format) = 0;

};

}

//---------------------------------------------------------------------------
#endif

