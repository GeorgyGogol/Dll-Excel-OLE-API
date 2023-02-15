//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelNameItem.h"
#include "uFunctions.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelNameItem::TExcelNameItem(TExcelObject* pParent, const Variant& data)
	: TExcelObjectTemplate<TExcelNameItem>(pParent, data)
{
}

TExcelNameItem::TExcelNameItem(const TExcelNameItem& src)
    : TExcelObjectTemplate<TExcelNameItem>(src)
{
}

TExcelNameItem::~TExcelNameItem() 
{}

/// @param val Значение
/// @note Передаваемое значение должно записаться именно как значение
TExcelNameItem* TExcelNameItem::SetValue(const Variant& val)
{
	CorrectInsert::InsertIntoVarCell(val, vData);
	return this;
}

/// @param val Строка
TExcelNameItem* TExcelNameItem::SetValue(const String& val)
{
	vData.OlePropertySet("Value", System::StringToOleStr(val));
	return this;
}

/// @param address Формула или ссылка на ячейки(у), куда будет указывать это имя
TExcelNameItem* TExcelNameItem::SetRefers(const String& address)
{
	String itemValue = "=\"" + StringReplace(address, "\"", "\"\"", TReplaceFlags()<<rfReplaceAll) + "\"";
	vData.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
	return this;
}

}
