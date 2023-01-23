//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelWorkbook.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
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

TExcelSheet* TExcelWorkbook::GetSheet(const String& sheetName) {
	seekAndSetDataChild("Sheets", sheetName);
	TExcelSheet* out = new TExcelSheet(this, vDataChild);
	return out;
}

TExcelSheet* TExcelWorkbook::GetSheet(unsigned int N) {
    seekAndSetDataChild("Sheets", (int)N);
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

TExcelWorkbook* TExcelWorkbook::Save()
{
	vData.OleProcedure("Save");
}

TExcelWorkbook* TExcelWorkbook::Save(const String filePath)
{
	if (filePath.Length() > 0){
		vData.OleProcedure("SaveAs", System::StringToOleStr(filePath));
	}
	else {
		vData.OleProcedure("Save");
	}
}

TExcelNameItem* TExcelWorkbook::GetNameItem(const String& itemName)
{
    vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", System::StringToOleStr(itemName));
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

TExcelNameItem* TExcelWorkbook::GetNameItem(unsigned int N)
{
    seekAndSetDataChild("Names", N);
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

TExcelNameItem* TExcelWorkbook::AddNamedItem(const String& itemName)
{
	vDataChild = vData.OlePropertyGet("Names");
	vDataChild.OleProcedure("Add", System::StringToOleStr(itemName));
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

}


