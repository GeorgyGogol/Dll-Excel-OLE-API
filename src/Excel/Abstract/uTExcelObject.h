#ifndef uTExcelObjectH
#define uTExcelObjectH

#include "uDll.h"
#include "uTExcelObjectNode.h"
#include "uTExcelObjectData.h"
//---------------------------------------------------------------------------
namespace exl {
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

	TExcelObject* GetCurrentSelectedChild(); ///< Какой-то общий метод

	TExcelObject* Show(); ///< Скрыть
	TExcelObject* Hide(); ///< Отобразить

	TExcelObject* SetName(const String& newName); ///< Установить имя
	String GetName(); ///< Получить имя в форме String

#ifdef ENABLE_USAGE_STATISTIC
	unsigned int SizeOfThis(); ///< Получить размер объекта в байтах
#endif
};

}
/// @}
#endif

