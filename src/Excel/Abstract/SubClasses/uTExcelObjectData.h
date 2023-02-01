#ifndef uTExcelObjectDataH
#define uTExcelObjectDataH

#include "uDll.h"
//---------------------------------------------------------------------------
namespace exl {
/// @addtogroup Abstract
/// @{
/**
 * @brief Класс-хранилище для Variant
 * 
 * Является абстрактным понятием.
 * 
 * Зона ответственности: хранение основной варианты объекта и "буферной"
 * дочерней. Также то, есть ли что в основной варианте или нет - его
 * задача, которая решается методом checkDataValide().
 * 
 * Да, по сути дочерняя варианта является неким подобием "глобальной 
 * переменной", доступной для всего дочернего класса. Может быть излишним,
 * но это позволяет обратиться к последнему выбранному объекту. На сколько
 * подход хорош - вопрос открытый.
 * 
 * Типовые решения взаимодействия также хранятся здесь.
 * 
 * Не имеет открытых методов, так как является чисто служебным элементом.
 */
class DLL_EI TExcelObjectData
{
public:
    TExcelObjectData();
    TExcelObjectData(const Variant& data);
	TExcelObjectData(const TExcelObjectData&);
protected:
	~TExcelObjectData();
    //void operator=();

protected:
    Variant vData; ///< Основная переменная
    Variant vDataChild; ///< Переменная дочернего элемента

    void checkDataValide(); ///< @brief Проверка варианты на корректность

    // Обще-типовые обертки функций для внутреннего использования
	unsigned int getChildCountByType(const String& oType); ///< Посчитать кол-во дочерних элементов
	void seekAndSetDataChild(const String& oType, const String& name); ///< Найти и установить дочернюю варианту (по имени)
    void seekAndSetDataChild(const String& oType, unsigned int Num); ///< Найти и установить дочернюю варианту (по номеру)

public:
    /// Это для особых извращений, которых нет в типовых решениях
	Variant getVariant();

};

}
/*! @} */
#endif

