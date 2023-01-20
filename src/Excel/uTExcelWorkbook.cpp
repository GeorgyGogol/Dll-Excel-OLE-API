//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelWorkbook.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {

TExcelWorkbook::TExcelWorkbook(const TExcelWorkbook& src)
	: TExcelObjectTemplate<TExcelWorkbook>(src)
{}

TExcelWorkbook::TExcelWorkbook(TExcelObject* pParent, const Variant& data)
    : TExcelObjectTemplate<TExcelWorkbook>(pParent, data)
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
	TExcelSheet* out = new TExcelSheet(this, vDataChild);
	return out;
}

TExcelSheet* TExcelWorkbook::SelectSheet(unsigned int N) {
    seekAndSetDataChild("Sheets", (int)N);
	vDataChild.OleProcedure("Activate");
	TExcelSheet* out = new TExcelSheet(this, vDataChild);
	return out;
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, const String &sheetName, const String &tableTitle, const String &tableName, bool needDisableSet)
{
	//TExcelSheet* aim = CreateSheet(sheetName);
	//TExcelTable* out;
    return CreateSheet(sheetName)->CreateTable(1, 1, dataSet, tableTitle, tableName, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, const String &sheetName, const String &tableTitle, bool needDisableSet)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, dataSet, tableTitle, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, const String &sheetName, bool needDisableSet)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, dataSet, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, bool needDisableSet)
{
    return CreateSheet()->CreateTable(1, 1, dataSet, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, const String &sheetName, const String &tableTitle, const String &tableName, bool needDisableSet)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, gridEh, tableTitle, tableName, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, const String &sheetName, const String &tableTitle, bool needDisableSet)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, gridEh, tableTitle, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, const String &sheetName, bool needDisableSet)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, gridEh, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, bool needDisableSet)
{
    return CreateSheet()->CreateTable(1, 1, gridEh);
}

}

