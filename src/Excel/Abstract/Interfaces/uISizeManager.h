#ifndef UISIZEMANAGER_H
#define UISIZEMANAGER_H
/** @file
 * @brief Содержит Интерфейс настройки размеров
 */

//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Интерфейс настройки размеров ячейки/ячеек
 * @ingroup Interfaces
 * 
 * Отвечает за добавление методов настройки размеров ячейки или ячеек
 * 
 * ---------------------------------------------------------------------- **/
template<class T>
class ISizeManager
{
	/// @brief Получить интерфейс
	/// @return Интерфейс текущего объекта
	virtual ISizeManager<T>* GetSizeInterface() = 0;


	virtual T* SetWidth(unsigned int points) = 0;
	// OlePropertySet("RowHeight", points);

	virtual T* SetHeight() = 0;
	
	//virtual T* AutoSize() = 0;
};

}

//---------------------------------------------------------------------------
#endif // UISIZEMANAGER_H
