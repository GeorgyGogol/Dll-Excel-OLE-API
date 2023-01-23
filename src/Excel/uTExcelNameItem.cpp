//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelNameItem.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
    //TExcelNameItem::TExcelNameItem();

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

TExcelNameItem* TExcelNameItem::SetValue(const Variant& val)
{
	//vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", System::StringToOleStr(itemName));
	//Variant vbuf;


	InsertIntoSingleVariant(val, vData);
	//vData.OlePropertySet("Value", val);
	return this;
}

TExcelNameItem* TExcelNameItem::SetValue(const String& val)
{
	vData.OlePropertySet("Value", System::StringToOleStr(val));
	return this;
}

TExcelNameItem* TExcelNameItem::SetRefers(const String& address)
{
	String itemValue = "=\"" + StringReplace(address, "\"", "\"\"", TReplaceFlags()<<rfReplaceAll) + "\"";
	vData.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
	return this;
}

}
