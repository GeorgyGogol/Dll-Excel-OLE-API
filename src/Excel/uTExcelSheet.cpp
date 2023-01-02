//---------------------------------------------------------------------------

#pragma hdrstop

#include "uTExcelSheet.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelSheet::TExcelSheet(TExcelObject* pParent, const Variant& data) 
    : TExcelObjectRanged(pParent, data)
{
	vDataChild = vData.OlePropertyGet("Range", "A1");
}

TExcelSheet::TExcelSheet(const TExcelSheet& src)
	: TExcelObjectRanged(src)
{
}

TExcelSheet::~TExcelSheet() {
}

TExcelCell* TExcelSheet::SelectCell(unsigned int col, unsigned int row) {
    return (TExcelCell*)select(col, row);
}

TExcelCells* TExcelSheet::SelectCells(
    unsigned int startColumn, unsigned int startRow,
    unsigned int endColumn, unsigned int endRow
    )
{
    return (TExcelCells*)selectRange(startColumn, startRow, endColumn, endRow);
}


TExcelCells* TExcelSheet::SelectColumn(unsigned int column) {
	//AnsiString currRange;
	//currRange.printf("%s:%s", ColToStrA(column).c_str(), ColToStrA(column).c_str());
    //return (TExcelCells*)selectRange(currRange);
	return 0;
}

//void TExcelSheet::HideColumn(unsigned int column) {
//	AnsiString currRange;
//    currRange.printf("%s:%s", ColToStrA(column).c_str(), ColToStrA(column).c_str());
//    vData.OlePropertyGet("Range", System::StringToOleStr(currRange)).OleProcedure("Hide");
//}

TExcelCells* TExcelSheet::SelectRow(unsigned int row) {
	//AnsiString currRange;
    //currRange.printf("%i:%i", row, row);
    //return (TExcelCell*)selectRange(currRange);
	return 0;
}

TExcelCells* TExcelSheet::InsertDataSet(
	unsigned int startColumn, unsigned int startRow,
	TDataSet* dataSet,
	bool needInsertFieldNames,
	bool needDisableSet
	)
{
	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (needDisableSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, dataSet);
	TExcelCells* range = creator->InsertData(startColumn, startRow, needInsertFieldNames);
	delete creator;

	// Вернем GUI (при необходимости)
	dataSet->RecNo = userPos;
	if (needDisableSet) dataSet->EnableControls();

	return range;


	///////////////////////////////////////////////////	

/*
	// Системное - строки для хождения по листу, размер массива
	unsigned int excelRowPos = 1;
	unsigned int ArrSize;

	// Размеры таблицы
	unsigned int nRecords = dataSet->RecordCount;
	unsigned int nColumns = dataSet->Fields->Count;

	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (needDisableSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Создадим массив данных
	ArrSize = nRecords;
	if (needInsertFieldNames) nRecords++; // +1 если нужны заголовки
	Variant varArr(OPENARRAY(int, (1, ArrSize, 1, nColumns)), varVariant); // TODO: OPENARRY

	// Заполним массив
	// Если есть нужда - заголовки
	if (needInsertFieldNames){
		String ColumnName;
		for (unsigned int col = 0; col < nColumns; ++col){
			ColumnName = dataSet->Fields->Fields[col]->FieldName;
			varArr.PutElement(ColumnName, excelRowPos, col + 1);
		}
		excelRowPos++;
	}
	// И само содержимое
	for (dataSet->First(); !dataSet->Eof && (excelRowPos - 2) < nRecords; dataSet->Next(), ++excelRowPos) {
		for (unsigned int col = 0; col < nColumns; ++col) {
			varArr.PutElement(VarToStr(dataSet->Fields->Fields[col]->Value), excelRowPos, col + 1);
        }
    }

	// Вставим данные
	TExcelCells* range = SelectCells(startColumn, startRow, startColumn + nColumns - 1, startRow + nRecords);
	range->Insert(varArr);

	// Вернем GUI (при необходимости)
	dataSet->RecNo = userPos;
	if (needDisableSet) dataSet->EnableControls();

	return range;
	*/
}

TExcelTable* TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow,
	TDataSet* dataSet, const String& tableName, const String& tableTitle,
	bool needDisableSet
	)
{
	// Запомним GUI
	unsigned int userPos = dataSet->RecNo;
	if (needDisableSet) dataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, dataSet, tableName, tableTitle);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator;

	// Вернем GUI
	dataSet->RecNo = userPos;
	if (needDisableSet) dataSet->EnableControls();

	return table;
}

TExcelTable* TExcelSheet::CreateTable(
	unsigned int startColumn, unsigned int startRow,
	TDBGridEh* gridEh, const String& tableName, const String& tableTitle,
	bool needDisableSet
	)
{
	// Запомним GUI
	unsigned int userPos = gridEh->DataSource->DataSet->RecNo;
	if (needDisableSet) gridEh->DataSource->DataSet->DisableControls(); // По желанию, т.к. может быть ОЧЕНЬ НЕПРИЯТНАЯ ЛЕДИ-БАГ

	// Сделаем вид, что работаем
	TTableCreator* creator = new TTableCreator(this, gridEh, tableName, tableTitle);
	TExcelTable* table = creator->CreateTable(startColumn, startRow);
	delete creator; // Мавр сделал свое дело - мавр может уходить.

	// Вернем GUI
	gridEh->DataSource->DataSet->RecNo = userPos;
	if (needDisableSet) gridEh->DataSource->DataSet->EnableControls();

	return table;
}

TExcelTable* TExcelSheet::CreatePivotTable(
	unsigned int startColumn, unsigned int startRow,
	TExcelTable* srcTable,
	TPivotSettings* pivotSettings
	)
{
	TExcelTable* out;

	TPivotTableCreator* creator = new TPivotTableCreator(srcTable);
	out = creator->CreateTable(startColumn, startRow, pivotSettings);

	return out;
}


}


