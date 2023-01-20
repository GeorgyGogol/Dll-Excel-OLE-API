//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObjectData.h"
#include "uTExcelDataExceptions.h"

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

TExcelObjectData::TExcelObjectData(const TExcelObjectData& src)
{
    vData = src.vData;
	vDataChild = src.vDataChild;
}

TExcelObjectData::~TExcelObjectData()
{}


void TExcelObjectData::checkDataValide()
{
	if (vData.IsNull()) throw ExcelDataNullException();
}

unsigned int TExcelObjectData::getChildCountByType(const String& oType)
{
	checkDataValide();
	unsigned int out = vData.OlePropertyGet(System::StringToOleStr(oType)).OlePropertyGet("Count");
	return out;
}

void TExcelObjectData::seekAndSetDataChild(const String& oType, const String& name)
{
	checkDataValide();
	vDataChild = vData.OlePropertyGet(System::StringToOleStr(oType), System::StringToOleStr(name));
}

void TExcelObjectData::seekAndSetDataChild(const String& oType, unsigned int Num)
{
	checkDataValide();
    vDataChild = vData.OlePropertyGet(System::StringToOleStr(oType)).OlePropertyGet("Item", Num);
}

Variant TExcelObjectData::getVariant(){
    return vData;
}

}


