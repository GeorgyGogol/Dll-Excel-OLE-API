//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObjectData.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelObjectData::TExcelObjectData() 
    : vData(Null()), vDataChild(Null())
{

}

TExcelObjectData::TExcelObjectData(const Variant& data) 
    : vData(data), vDataChild(Null())
{

}

TExcelObjectData::TExcelObjectData(const TExcelObjectData& src) {
    vData = src.vData;
    vDataChild = src.vDataChild;
}

TExcelObjectData::~TExcelObjectData()
{}


unsigned int TExcelObjectData::getChildCountByType(const String& oType)
{
    unsigned int out = vData.OlePropertyGet(System::StringToOleStr(oType)).OlePropertyGet("Count");
	return out;
}

void TExcelObjectData::seekAndSetDataChild(const String& oType, const String& name) 
{
	vDataChild = vData.OlePropertyGet(System::StringToOleStr(oType), System::StringToOleStr(name));
}

void TExcelObjectData::seekAndSetDataChild(const String& oType, unsigned int Num)
{
    vDataChild = vData.OlePropertyGet(System::StringToOleStr(oType)).OlePropertyGet("Item", Num);
}

Variant TExcelObjectData::getVariant(){
    return vData;
}

void TExcelObjectData::Show()
{
    vData.OlePropertySet("Visible", true);
}

void TExcelObjectData::Hide()
{
    vData.OlePropertySet("Visible", false);
}

}

