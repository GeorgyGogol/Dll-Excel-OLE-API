#ifndef uTExcelObjectTemplateH
#define uTExcelObjectTemplateH

#include "uTExcelObject.h"
//---------------------------------------------------------------------------
namespace exl {
/** ------------------------------------------------------------------------- 
 * @brief Базовый шаблон для наследования дочерних классов
 * @ingroup Templates
 * 
 * Основной шаблон, используемый для наследования. При наследовании нового
 * дочернего класса необходимо добавить в cpp-файл предопределение и шаблон,
 * использующий новый класс. Это связано с тем, что реализация шаблонных
 * методов помещена в cpp-файл. 
 * Суть реализованных методов: вызвать метод родительский и скастовать this,
 * который возвращается, в дочерний тип.
 * 
 * ---------------------------------------------------------------------- **/
template<class T>
class DLL_EI TExcelObjectTemplate : public TExcelObject
{
public:
	TExcelObjectTemplate();
	TExcelObjectTemplate(TExcelObject* pParent, const Variant& data);
	TExcelObjectTemplate(const TExcelObject& src);
protected:
	~TExcelObjectTemplate();

protected:

public:
	T* Show();
	T* Hide();

	T* SetName(const String& newName);

};

}

//---------------------------------------------------------------------------
#endif

