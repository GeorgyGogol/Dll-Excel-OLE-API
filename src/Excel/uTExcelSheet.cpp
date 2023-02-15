//---------------------------------------------------------------------------

#pragma hdrstop

#include "uTExcelSheet.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelSheet::TExcelSheet(TExcelObject* pParent, const Variant& data) :
    TExcelObjectRangedTemplate<TExcelSheet>(pParent, data),
	ACreateTableController()
{
	vDataChild = vData.OlePropertyGet("Range", "A1");
}

TExcelSheet::TExcelSheet(const TExcelSheet& src) :
	TExcelObjectRangedTemplate<TExcelSheet>(src),
	ACreateTableController()
{
}

TExcelSheet::~TExcelSheet() {
}

TExcelCells* TExcelSheet::GetCell(unsigned int col, unsigned int row)
{
	setSingle(col, row);
	TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

TExcelCells* TExcelSheet::GetCell(
	unsigned int startColumn, unsigned int startRow, 
	unsigned int endColumn, unsigned int endRow
	)
{
	setRange(startColumn, startRow, endColumn, endRow);
	TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}


TExcelCells* TExcelSheet::SelectCell(unsigned int col, unsigned int row) {
	selectSingle(col, row);
	TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

TExcelCells* TExcelSheet::SelectCells(
    unsigned int startColumn, unsigned int startRow,
    unsigned int endColumn, unsigned int endRow
	)
{
	selectRange(startColumn, startRow, endColumn, endRow);
	TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

TExcelCells* TExcelSheet::SelectColumn(unsigned int column) {
	AnsiString currRange;
	currRange.printf("%s:%s", ColToStrA(column).c_str(), ColToStrA(column).c_str());
	seekAndSetDataChild("Range", currRange);
	TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelCells* TExcelSheet::SelectColumns(unsigned int colStart, unsigned int colEnd) {
	// TODO: Check colStart < colEnd
	AnsiString currRange;
	currRange.printf("%s:%s", ColToStrA(colStart).c_str(), ColToStrA(colEnd).c_str());
	seekAndSetDataChild("Range", currRange);
	TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelCells* TExcelSheet::SelectRow(unsigned int row) {
	AnsiString currRange;
    currRange.printf("%i:%i", row, row);
	seekAndSetDataChild("Range", currRange);
	TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelCells* TExcelSheet::SelectRows(unsigned int rowStart, unsigned int rowEnd) {
	AnsiString currRange;
    currRange.printf("%i:%i", rowStart, rowEnd);
	seekAndSetDataChild("Range", currRange);
	TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelCells* TExcelSheet::InsertDataSet(
	unsigned int startColumn, unsigned int startRow,
	TDataSet* dataSet,
	bool needInsertFieldNames
	)
{
	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (NeedDisableDataSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, dataSet);
	TExcelCells* range = creator->InsertData(startColumn, startRow, needInsertFieldNames);
	delete creator;

	// Вернем GUI (при необходимости)
	dataSet->RecNo = userPos;
	if (NeedDisableDataSet) dataSet->EnableControls();

	return range;
}

TExcelTable* TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow,
	TDataSet* dataSet, const String& tableTitle, const String& tableName
	)
{
	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (NeedDisableDataSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, dataSet, tableTitle, tableName);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator;

	// Вернем GUI
	dataSet->RecNo = userPos;
	if (NeedDisableDataSet) dataSet->EnableControls();

	return table;
}

TExcelTable *TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow, 
	TDataSet *dataSet, const String &tableTitle
	)
{
	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (NeedDisableDataSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, dataSet, tableTitle);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator;

	// Вернем GUI
	dataSet->RecNo = userPos;
	if (NeedDisableDataSet) dataSet->EnableControls();

	return table;
}


TExcelTable* TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow,
	TDataSet* dataSet
	)
{
	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (NeedDisableDataSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, dataSet);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator;

	// Вернем GUI
	dataSet->RecNo = userPos;
	if (NeedDisableDataSet) dataSet->EnableControls();

	return table;
}

TExcelTable* TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow,
	TDBGridEh* gridEh, const String& tableTitle, const String& tableName
	)
{
	// Запомним GUI
	unsigned int userPos = gridEh->DataSource->DataSet->RecNo;
	if (NeedDisableDataSet) gridEh->DataSource->DataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, gridEh, tableTitle, tableName);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator; // Мавр сделал свое дело - мавр может уходить.

	// Вернем GUI
	gridEh->DataSource->DataSet->RecNo = userPos;
	if (NeedDisableDataSet) gridEh->DataSource->DataSet->EnableControls();

	return table;
}

TExcelTable* TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow,
	TDBGridEh* gridEh, const String& tableTitle
	)
{
	// Запомним GUI
	unsigned int userPos = gridEh->DataSource->DataSet->RecNo;
	if (NeedDisableDataSet) gridEh->DataSource->DataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, gridEh, tableTitle);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator; // Мавр сделал свое дело - мавр может уходить.

	// Вернем GUI
	gridEh->DataSource->DataSet->RecNo = userPos;
	if (NeedDisableDataSet) gridEh->DataSource->DataSet->EnableControls();

	return table;
}

TExcelTable *TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow, 
	TDBGridEh *gridEh
	)
{
	// Запомним GUI
	unsigned int userPos = gridEh->DataSource->DataSet->RecNo;
	if (NeedDisableDataSet) gridEh->DataSource->DataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, gridEh);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator; // Мавр сделал свое дело - мавр может уходить.

	// Вернем GUI
	gridEh->DataSource->DataSet->RecNo = userPos;
	if (NeedDisableDataSet) gridEh->DataSource->DataSet->EnableControls();

	return table;
}

TExcelTable* TExcelSheet::GetTable(const String& tableName)
{
	vDataChild = vData.OlePropertyGet("ListObjects").OlePropertyGet("Item", System::StringToOleStr(tableName));

	TExcelTable* out = new TExcelTable(this, vDataChild);
	return out;
}

TExcelTable* TExcelSheet::CreatePivotTable(
	unsigned int startColumn, unsigned int startRow,
	TExcelTable* srcTable,
	TPivotSettings* pivotSettings
	)
{
	TExcelTable* out;

	TPivotTableCreator* creator = new TPivotTableCreator(srcTable);
	out = creator->CreateTable(this, startColumn, startRow, pivotSettings);

	return out;
}

TExcelNameItem* TExcelSheet::GetNameItem(const String& itemName)
{
    vDataChild = vData.OlePropertyGet("Names").OleFunction("Item", System::StringToOleStr(itemName));
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

TExcelNameItem* TExcelSheet::GetNameItem(unsigned int N)
{
    seekAndSetDataChild("Names", N);
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

TExcelNameItem* TExcelSheet::AddNamedItem(const String& itemName)
{
	vDataChild = vData.OlePropertyGet("Names");
	vDataChild.OleProcedure("Add", System::StringToOleStr(itemName));
    TExcelNameItem* out = new TExcelNameItem(this, vDataChild);
    return out;
}

}


