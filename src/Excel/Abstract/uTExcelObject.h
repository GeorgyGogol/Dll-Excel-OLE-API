#ifndef uTExcelObjectH
#define uTExcelObjectH

#include "uDll.h"
#include "uTExcelObjectNode.h"
#include "uTExcelObjectData.h"
//---------------------------------------------------------------------------
/// @brief Пространство, в котором располагаются объекты
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Базовый класс объектов
 * @ingroup Abstract
 * 
 * Абстракция, но которую уже можнно использовать в своих коварных целях.
 * 
 * Является точкой собрания абстрактных классов в один. Реализует все
 * максимально общие задачи.
 * 
 * Методы класса возвращают указатель на сам класс. Геттеры возвращают 
 * порожденные объекты.
 * @code
 * TExcelObject* oParent, oChild;
 * //<...>
 * oParent->SetName("NewName")->Show(); // Работа с методами
 * oChild = oParent->GetCurrentSelectedChild(); // Работа с Геттером
 * oChild = oParent->Hide(); // Ошибка: дочерний указатель указывает на сам объект
 * @endcode
 *
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelObject : public TExcelObjectNode, public TExcelObjectData
{
public:
	TExcelObject();
	TExcelObject(TExcelObject* pParent, const Variant& data);
	TExcelObject(const TExcelObject& src);
	virtual ~TExcelObject(); // ???

private:

protected:

public:
	TExcelObject* GetParent() const; ///< Получить родителя
	Variant GetParentVariant(); ///< Получить варианту родителя

	TExcelObject* Show(); ///< Скрыть объект
	TExcelObject* Hide(); ///< Отобразить объект

	TExcelObject* SetName(const String& newName); ///< Установить имя
	String GetName(); ///< Получить имя в форме String

	unsigned int SizeOfThis(); ///< Получить размер объекта в байтах
};

}

//---------------------------------------------------------------------------
#endif

