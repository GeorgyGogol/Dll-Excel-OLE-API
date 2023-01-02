//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelWorkbook.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {

TExcelWorkbook::TExcelWorkbook(const TExcelWorkbook& src)
    : TExcelObject(src)
{}

TExcelWorkbook::TExcelWorkbook(TExcelObject* pParent, const Variant& data)
    : TExcelObject(pParent, data)
{
	seekAndSetDataChild("Sheets", 1);
	vDataChild.OleProcedure("Activate");
}

TExcelWorkbook::~TExcelWorkbook()
{}

/* TExcelApp* TExcelWorkbook::getExcel() {
    return (TExcelApp*)getParent();
} */

unsigned int TExcelWorkbook::SheetCount() {
    return getChildCountByType("Sheets");
}

TExcelSheet* TExcelWorkbook::CreateSheet() {
	vDataChild = vData.OlePropertyGet("Sheets").OleFunction("Add");
    vDataChild.OleProcedure("Activate");
	return GetCurrentSheet();
}

TExcelSheet* TExcelWorkbook::CreateSheet(const String& sheetName) {
    TExcelSheet* out = CreateSheet();
	if (out){
		//out->setName(workbookName);
	}
	return out;
}

TExcelSheet* TExcelWorkbook::GetCurrentSheet() {
    TExcelSheet* out = 0;
	if (vDataChild != Null()){
		out = new TExcelSheet(this, vDataChild);	
	}
	return out;
}

TExcelSheet* TExcelWorkbook::SelectSheet(const String& sheetName) {
	seekAndSetDataChild("Sheets", sheetName);
	vDataChild.OleProcedure("Activate");
	return (TExcelSheet*)getCurrentSelectedChild();
}

TExcelSheet* TExcelWorkbook::SelectSheet(unsigned int N) {
    seekAndSetDataChild("Sheets", (int)N);
	vDataChild.OleProcedure("Activate");
	return (TExcelSheet*)getCurrentSelectedChild();
}

}

