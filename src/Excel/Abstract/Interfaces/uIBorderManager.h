#ifndef UIBORDERMANAGER_H
#define UIBORDERMANAGER_H
/** @file
 * @brief Содержит Интерфейс работы с грницами
 */

#include "uEnums.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Интерфейс для настройки грниц ячеек
 * @ingroup Interfaces
 * 
 * Отвечает за добавление методов настройки границ ячейки или ячеек
 * 
 * ---------------------------------------------------------------------- **/
template<class T>
class IBorderManager
{
public:
	/// @brief Получить интерфейс
	/// @return Интерфейс текущего объекта
	virtual IBorderManager<T>* GetBorderInterface() = 0;

	//virtual T* SetBordersAll(XlLineStyle style, XlBorderWeight weight) = 0;
	virtual T* SetBorder(XlBordersIndex border, XlLineStyle style, XlBorderWeight weight) = 0;

	//virtual T* RemoveBordersAll() = 0;
	virtual T* RemoveBorder(XlBordersIndex border) = 0;

	//virtual T* SetBorderLeft() = 0;
	//virtual T* SetBorderRight() = 0;
	//virtual T* SetBorderDown() = 0;
	//virtual T* SetBorderUp() = 0;
};

}

//---------------------------------------------------------------------------
#endif // UIBORDERMANAGER_H
