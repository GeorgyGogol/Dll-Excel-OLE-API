//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObject.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelObject::TExcelObject()
	: TExcelObjectNode(), TExcelObjectData()
{
}

TExcelObject::TExcelObject(TExcelObject* pParent, const Variant& data)
    : TExcelObjectNode(pParent), TExcelObjectData(data)
{
}

TExcelObject::TExcelObject(const TExcelObject& src)
	: TExcelObjectNode(src), TExcelObjectData(src)
{
}

TExcelObject::~TExcelObject()
{
}

/// @details По сути идет обращение к защищенному методу и приведение к
/// текущему классу
/// @return Указатель на родительский объект (+ приведенный к нужному типу)
TExcelObject* TExcelObject::GetParent() const
{
    return (TExcelObject*)getParentNode();
}

/// @details Берет родителя и если такой есть, обращается к его методу
/// @return Варианта родителя или Null(), если такого нет
Variant TExcelObject::GetParentVariant() 
{
    TExcelObject* dParent = GetParent();
    if (dParent)
        return dParent->getVariant();
    else return Null();
}

TExcelObject* TExcelObject::GetCurrentSelectedChild()
{
    TExcelObject* out = new TExcelObject(this, vDataChild);
    return out;
}

TExcelObject* TExcelObject::Show() {
    checkDataValide();
    vData.OlePropertySet("Visible", true);
    return this;
}

TExcelObject* TExcelObject::Hide() {
    checkDataValide();
    vData.OlePropertySet("Visible", false);
    return this;
}

TExcelObject* TExcelObject::SetName(const String& newName) {
    checkDataValide();
    vData.OlePropertySet("Name", System::StringToOleStr(newName));
    return this;
}

String TExcelObject::GetName()
{
    checkDataValide();
    Variant name = vData.OlePropertyGet("Name");
    return VarToStr(name);
}

#ifdef ENABLE_USAGE_STATISTIC
/// @details Метод обращается ко всем своим "дочкам" и спрашивает их размер
/// В свою очередь дочерние элементы считают размеры своих дочерних и возвращают
/// ответ инициатору. 
/// @bug Размеры подсчитываются не самих элементов, а их базового класса. Т.е.
/// фактический размер может быть чуть больше, если у дочерних классов есть свои,
/// частные, свойства.
/// @return Размер занимаемой памяти структурой
unsigned int TExcelObject::SizeOfThis() {
	unsigned int s = sizeof(*this);

	std::list<TExcelObjectNode*>::iterator it = Begin();

	for (it; it != End(); ++it) {
		s += static_cast<TExcelObject*>(*it)->SizeOfThis();
	}

	return s;
}
#endif

/// @brief Вставка варианты данных в варианту ячейки
/// @param vData Данные (массив OPENARRAY или строка)
/// @param vCell Переменная с Ole-объектом
/// @param sNullValue Значение по умолчанию
/// @todo Прописать логику и вынести в отдельный объект
void InsertIntoSingleVariant(const Variant& vData, Variant& vCell, const String& sNullValue)
{
	/*
	String buf;
	long double dBuf;
	TDateTime dtBuf;

	buf = VarToStrDef(vData, sNullValue.c_str());

	if(TryStrToFloat(buf, dBuf)){
		vCell.OlePropertySet("Value", dBuf);
	}
	else if (TryStrToTime(buf, dtBuf)){
		vCell.OlePropertySet("Value", dtBuf);
	}
	else {

	}
	*/

	vCell.OlePropertySet("Value", vData);
}

/// @brief Я вообще хз зачем, см. комментарии к InsertIntoSingleVariant()
void InsertIntoVarArray(const Variant& vData, Variant& varArr, unsigned int row, unsigned int col, const String& sNullValue)
{
	String buf;
	long double dBuf;
	TDateTime dtBuf;

	buf = VarToStrDef(vData, sNullValue.c_str());

	if(TryStrToFloat(buf, dBuf)){
		varArr.PutElement(dBuf, row, col);
	}
	else if (TryStrToTime(buf, dtBuf)){
		varArr.PutElement(dtBuf, row, col);
	}
	else {
		varArr.PutElement(buf, row, col);
	}
}

}


