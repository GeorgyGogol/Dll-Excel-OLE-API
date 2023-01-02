//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTTableCreator.h"
#include "uTExcelExceptions.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl{
//---------------------------------------------------------------------------    
TTableCreator::TTableCreator()
{}

TTableCreator::TTableCreator(TExcelObjectRanged* sheet, TDataSet* dataSet, const String& tableTitle, const String& tableName)
	: Sheet(NULL), Headers(NULL)
{
	PrepareNewData(sheet, dataSet, tableTitle, tableName);
}

TTableCreator::TTableCreator(TExcelObjectRanged* sheet, TDataSet* dataSet, const String& tableTitle)
	: Sheet(NULL), Headers(NULL)
{
	PrepareNewData(sheet, dataSet, tableTitle);
}

TTableCreator::TTableCreator(TExcelObjectRanged* sheet, TDataSet* dataSet)
	: Sheet(NULL), Headers(NULL)
{
	PrepareNewData(sheet, dataSet);
}

TTableCreator::TTableCreator(TExcelObjectRanged* sheet, TDBGridEh* gridEh, const String& tableName, const String& tableTitle)
	: Sheet(sheet), TableName(tableName), Title(tableTitle)
{
    init(gridEh);
}

TTableCreator::~TTableCreator()
{
	ResetData();
}

void TTableCreator::readDataSet(TDataSet* dataSet) {
	if (!Headers) throw ExcelTableCreatorException("readDataSet", "Headers is not created, cannot read DataSet!");

	// Размеры таблицы
	nRecords = dataSet->RecordCount;
	unsigned int cols = Headers->CountVisible();

	// Дальше подготовим данные к вставке
	// Поймем что как
	Variant varArr(OPENARRAY(int, (1, nRecords, 1, cols)), varVariant);

    // Считаем данные
	unsigned int tabColPassed; // Кол-во пропущенных столбиков
	unsigned int pos = 1; // Счетчик по записям

	// Буферы для разных типов
	String buf;
	long double dBuf;
	TDateTime dtBuf;

	for (dataSet->First(); !dataSet->Eof && pos < nRecords + 1; dataSet->Next(), pos++)
	{
        tabColPassed = 0;
		for (unsigned int tabCol = 0; tabCol < Headers->Count(); tabCol++) {
			if (!Headers->GetHeader(tabCol).Visible) {
				tabColPassed++;
				continue;
			}

			//if (!src->Fields[tabCol]->IsNull)
			buf = VarToStrDef(dataSet->Fields->Fields[tabCol]->Value, "0");
			//else buf = "";

			if(TryStrToFloat(buf, dBuf)){
				varArr.PutElement(dBuf, pos, tabCol + 1 - tabColPassed);
			}
			else if (TryStrToTime(buf, dtBuf)){
				varArr.PutElement(dtBuf, pos, tabCol + 1 - tabColPassed);
			}
			else {
				varArr.PutElement(buf, pos, tabCol + 1 - tabColPassed);
            }
        }
	}

	varData = varArr;
}

void TTableCreator::init(TDataSet* dataSet) {
	Headers = new TGridHeaders(dataSet);
	readDataSet(dataSet);
}

void TTableCreator::init(TDBGridEh* gridEh) {
	Headers = new TGridHeaders(gridEh);
	readDataSet(gridEh->DataSource->DataSet);
}

void TTableCreator::check(unsigned int& col, unsigned int& row)
{
	if (col < 1) col = 1;
	if (row < 1) row = 1;
}

void TTableCreator::ResetData() {
    Sheet = NULL;

    if (Headers) delete Headers;
	Headers = NULL;

	varData = Null();
	nRecords = 0;

	Title = "TableStyleLight1";
}

void TTableCreator::PrepareNewData(TExcelObjectRanged* sheet, TDataSet* dataSet, const String& tableTitle, const String& tableName)
{
	ResetData();

	Sheet = sheet;
	init(dataSet);
	TableName = tableName;
	Title = tableTitle;
}

void TTableCreator::PrepareNewData(TExcelObjectRanged* sheet, TDBGridEh* gridEh, const String& tableTitle, const String& tableName) {
	ResetData();

	Sheet = sheet;
	init(gridEh);
	TableName = tableName;
	Title = tableTitle;
}

bool TTableCreator::CanCreate() const
{
	return Sheet && Headers;
}

TExcelCells* TTableCreator::InsertData(unsigned int col, unsigned int row, bool needInsertFieldNames)
{
	if (!CanCreate()) throw ExcelTableCreatorException("CreateTable", 
		"Cannot create Table because there is not init Headers and/or information about aimed sheet not defined!");
	
	check(col, row);

	unsigned int sCol = col, sRow = row;

	if (needInsertFieldNames){
		TExcelCells* TableHeadersCells;
		TableHeadersCells = (TExcelCells*)Sheet->selectRange(col, row, col + Headers->CountVisible() - 1, row + nRecords + Headers->Deep() - 1);
		TableHeadersCells->Insert(Headers->generateVariant());
		sRow = row + nRecords + Headers->Deep();
		delete TableHeadersCells;
	}

	TExcelCells* TableDataCells;
	TableDataCells = (TExcelCells*)Sheet->selectRange(sCol, sRow, sCol + Headers->CountVisible() - 1, sRow + nRecords - 1);
	TableDataCells->Insert(varData);
	delete TableDataCells;

	TExcelCells* out;
	out = (TExcelCells*)Sheet->selectRange(col, row, sCol + Headers->CountVisible(), sRow + nRecords - 1);
	out->Insert(varData);
	return out;
}

TExcelTable* TTableCreator::CreateTable(unsigned int col, unsigned int row) {
	if (!CanCreate()) throw ExcelTableCreatorException("CreateTable", 
		"Cannot create Table because there is not init Headers and/or information about aimed sheet not defined!");
	if (Headers->Deep() != 1) throw ExcelTableCreatorException("CreateTable", 
		"Cannot create Table because there is more than one Header in depth!");

	check(col, row);
	
	unsigned int maxColumnOnSheet = col + Headers->CountVisible();

	TExcelCells* TableTitleCells;
	TExcelTableHeaders* tableHeaders;
	TExcelCells* TableDataCells;

    // Если есть заголовок - вставим
	if (Title.Length() > 0) {
		TableTitleCells = (TExcelCells*)(Sheet->selectRange(col, row, maxColumnOnSheet, row));
		TableTitleCells->Merge()->InsertString(Title);
		row += 2;
    }

	// Вставим Заголовки
	TExcelCells* TableHeadersCells;
	TableHeadersCells = (TExcelCells*)Sheet->selectRange(col, row, maxColumnOnSheet, row);
	TableHeadersCells->Insert(Headers->generateVariant());
	tableHeaders = new TExcelTableHeaders(TableHeadersCells);
	delete TableHeadersCells;
    
    // Вставим содержимое
	TableDataCells = (TExcelCells*)Sheet->selectRange(
            col, 
			row + 1,
			maxColumnOnSheet,
			row + 1 + nRecords - 1 
			);
	// Объясняю свой гений:
	// Да, смещение на 1 вниз - заголовок. Потом вниз идем на nRecords - 1, т.к. поправка
	// на 1 (для размещения 1й записи на 2й строке: строка 2 + 1 строка заголовка (это 3)
	// + 1 запись (4, тут уже выделено 2 строки вниз) - 1 = 3 => после 2й строки (заголовки)
	// будет выделена 3я строка)
	TableDataCells->Insert(varData);
	TExcelCells* tableData = new TExcelCells(Sheet, varData);

    // Создадим таблицу в Экселе
	Sheet->selectRange(col, row, maxColumnOnSheet, row + 1 + nRecords - 1);
	
	Variant vTable = Sheet->getVariant().OlePropertyGet("ListObjects").OleFunction("Add");
	
	if (TableName.Length() > 0) 
		vTable.OlePropertySet("Name", System::StringToOleStr(TableName));
	else
		TableName = VarToStr(vTable.OlePropertyGet("Name"));
	
	if (TableStyle.Length() > 0) vTable.OlePropertySet("TableStyle", TableStyle);

    // А теперь сделаем объект С++
	TExcelTable* table = new TExcelTable(Sheet, TableName, tableHeaders, TableDataCells, TableTitleCells);

    return table;
}


}

