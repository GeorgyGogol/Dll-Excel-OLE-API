//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelWorkbook.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelWorkbook::TExcelWorkbook(TExcelObject* pParent, const Variant& data) :
	TExcelObjectTemplate<TExcelWorkbook>(pParent, data)
	//ACreateTableController()
{
	seekAndSetDataChild("Sheets", 1);
	vDataChild.OleProcedure("Activate");
}

TExcelWorkbook::TExcelWorkbook(const TExcelWorkbook& src) :
	TExcelObjectTemplate<TExcelWorkbook>(src)
	//ACreateTableController()
{}

TExcelWorkbook::~TExcelWorkbook()
{}

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

TExcelWorkbook* TExcelWorkbook::Save()
{
	vData.OleProcedure("Save");
	return this;
}

TExcelWorkbook* TExcelWorkbook::Save(const String filePath)
{
	if (filePath.Length() > 0){
		vData.OleProcedure("SaveAs", System::StringToOleStr(filePath));
	}
	else {
		vData.OleProcedure("Save");
	}
    return this;
}

/* TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, const String &sheetName, const String &tableTitle, const String &tableName)
{
	//TExcelSheet* aim = CreateSheet(sheetName);
	//TExcelTable* out;
    return CreateSheet(sheetName)->CreateTable(1, 1, dataSet, tableTitle, tableName, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, const String &sheetName, const String &tableTitle)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, dataSet, tableTitle, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet, const String &sheetName)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, dataSet, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDataSet *dataSet)
{
    return CreateSheet()->CreateTable(1, 1, dataSet, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, const String &sheetName, const String &tableTitle, const String &tableName)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, gridEh, tableTitle, tableName, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, const String &sheetName, const String &tableTitle)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, gridEh, tableTitle, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh, const String &sheetName)
{
    return CreateSheet(sheetName)->CreateTable(1, 1, gridEh, needDisableSet);
}

TExcelTable *TExcelWorkbook::CreateTable(TDBGridEh *gridEh)
{
    return CreateSheet()->CreateTable(1, 1, gridEh);
} */

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

