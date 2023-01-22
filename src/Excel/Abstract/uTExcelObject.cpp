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

TExcelObject* TExcelObject::GetParent() const 
{
    return (TExcelObject*)getParentNode();
}

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
/*
TExcelNameItem* TExcelObject::GetNameItem(const String& itemName)
{
    vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", System::StringToOleStr(itemName));
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

TExcelNameItem* TExcelObject::GetNameItem(unsigned int N)
{
    seekAndSetDataChild("Names", N);
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

TExcelNameItem* TExcelObject::AddNamedItem(const String& itemName)
{
	vDataChild = vData.OlePropertyGet("Names");
	vDataChild.OleProcedure("Add", System::StringToOleStr(itemName));
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
	return out;
}       */

/* String StrToFormul(const String& str){
    return "=\"" + str + "\"";
} */

/* void* TExcelObject::AddNamedItem(const String& itemName, const String& formula, const String& comment)
{
    checkDataValide();
    String itemValue = "=\"" + formula + "\"";
	vDataChild = vData.OlePropertyGet("Names");
	vDataChild.OleProcedure("Add", System::StringToOleStr(itemName), System::StringToOleStr(itemValue));
	vDataChild.OlePropertySet("Comment", System::StringToOleStr(comment));
	//return GetCurrentSelectedChild();
} */

/* void* TExcelObject::SetItem(const String& itemName, const String& data)
{
    checkDataValide();
    String itemValue = "=\"" + data + "\"";
    vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", System::StringToOleStr(itemName));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}

void* TExcelObject::SetItem(const String& itemName, const Variant& data)
{
	checkDataValide();
	String itemValue = "=\"" + VarToStr(data) + "\"";
	vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", System::StringToOleStr(itemName));
	vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}

void* TExcelObject::SetItem(unsigned int i, const String& data)
{
    checkDataValide();
    String itemValue = "=\"" + data + "\"";
	vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", int(i));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}

void* TExcelObject::SetItem(unsigned int i, const Variant& data)
{
    checkDataValide();
    String itemValue = "=\"" + VarToStr(data) + "\"";
    vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", int(i));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
} */


String TExcelObject::GetName()
{
    checkDataValide();
    Variant name = vData.OlePropertyGet("Name");
    return VarToStr(name);
}

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


