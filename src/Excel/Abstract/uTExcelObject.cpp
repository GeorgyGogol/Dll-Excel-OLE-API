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

void* TExcelObject::AddNamedItem(const String& itemName)
{
    checkDataValide();
    vDataChild = vData.OlePropertyGet("Names").OleProcedure("Add", System::StringToOleStr(itemName), "=\"\"");
    //return GetCurrentSelectedChild();
}

void* TExcelObject::AddNamedItem(const String& itemName, const String& formula)
{
    checkDataValide();
    String itemValue = "=\"" + formula + "\"";
    vDataChild = vData.OlePropertyGet("Names").OleProcedure("Add", System::StringToOleStr(itemName), System::StringToOleStr(itemValue));
    //return GetCurrentSelectedChild();
}

/* String StrToFormul(const String& str){
    return "=\"" + str + "\"";
} */

void* TExcelObject::AddNamedItem(const String& itemName, const String& formula, const String& comment)
{
    checkDataValide();
    String itemValue = "=\"" + formula + "\"";
    vDataChild = vData.OlePropertyGet("Names").OleProcedure("Add", System::StringToOleStr(itemName), System::StringToOleStr(itemValue));
    vDataChild.OlePropertySet("Comment", System::StringToOleStr(comment));
    //return GetCurrentSelectedChild();
}

void* TExcelObject::SetItem(const String& itemName, const String& data)
{
    checkDataValide();
    String itemValue = "=\"" + data + "\"";
    vDataChild = vData.OlePropertyGet("Names").OlePropertyGet("Item", System::StringToOleStr(itemName));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}

void* TExcelObject::SetItem(const String& itemName, const Variant& data)
{
    checkDataValide();
    String itemValue = "=\"" + VarToStr(data) + "\"";
    vDataChild = vData.OlePropertyGet("Names").OlePropertyGet("Item", System::StringToOleStr(itemName));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}

void* TExcelObject::SetItem(unsigned int i, const String& data)
{
    checkDataValide();
    String itemValue = "=\"" + data + "\"";
    vDataChild = vData.OlePropertyGet("Names").OlePropertyGet("Item", int(i));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}

void* TExcelObject::SetItem(unsigned int i, const Variant& data)
{
    checkDataValide();
    String itemValue = "=\"" + VarToStr(data) + "\"";
    vDataChild = vData.OlePropertyGet("Names").OlePropertyGet("Item", int(i));
    vDataChild.OlePropertySet("RefersTo", System::StringToOleStr(itemValue));
}


String TExcelObject::GetName()
{
    checkDataValide();
    Variant name = vData.OlePropertyGet("Name");
    String out = VarToStr(name);
    return out;
}

}

