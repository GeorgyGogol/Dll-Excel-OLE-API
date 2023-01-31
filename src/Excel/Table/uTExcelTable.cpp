//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTable.h"
#include "uVariantCorrectInserter.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelTable::TExcelTable(TExcelObject* pSheet, const String& tableName)
    : TExcelObjectTemplate<TExcelTable>(pSheet, Null())
{
    Variant vParent = GetParentVariant();
    vData = vParent.OlePropertyGet("ListObjects").OlePropertyGet("Item", System::StringToOleStr(tableName));
}


TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable)
    : TExcelObjectTemplate<TExcelTable>(pSheet, vTable)
{
}

TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable, const Variant& vTitle)
    : TExcelObjectTemplate<TExcelTable>(pSheet, vTable)
{
    Title = new TExcelCells(this, vTitle);
}


TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable, TExcelCells* eTitle)
	: TExcelObjectTemplate<TExcelTable>(pSheet, vTable)
{
	//Variant vTitleData = eTitle->getVariant();
	//Title = new TExcelCells(this, vTitleData);
}

/*
TExcelTable::TExcelTable(TExcelObjectRanged* Sheet, const Variant& name, const Variant& headers, const Variant& data)
	: TExcelObjectRanged(*Sheet)
{
	NameCells = new TExcelCell(Sheet, name);
	Headers = new TExcelTableHeaders(Sheet, headers);
	Data = new TExcelTableData(Sheet, data);
}
*/

/*
TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells)
	: TExcelObjectTemplate<TExcelTable>(pSheet, vTable)
		//, Headers(headers), Data(data)
        , Title(titleCells)
{
}

TExcelTable::TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells)
	: TExcelObjectTemplate<TExcelTable>(pSheet, Null())
		//, Headers(headers), Data(data)
        , Title(titleCells)
{
    //vData = 
}

TExcelTable::TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data)
	: TExcelObjectTemplate<TExcelTable>(pSheet, Null())
        //, Headers(headers), Data(data)
        , Title(NULL)
{
}

/*
TExcelTable::TExcelTable(TExcelObjectRanged* Sheet, TExcelCell* cells, unsigned int headersCount)
    : TExcelObjectRanged(*Sheet)
{
    if (headersCount == 0) headersCount = 1;
    unsigned int columns = cells->GetColumnsCount();
    Variant vHeaders = cells->GetCellsFromCells(1, 1, columns, headersCount)->getVariant();
    Variant vData = cells->GetCellsFromCells(1, headersCount + 1, columns, cells->GetRowCount() - headersCount)->getVariant();

    Headers = new TExcelTableHeaders(Sheet, vHeaders);
    Data = new TExcelTableData(Sheet, vData);
}
*/

TExcelTable::~TExcelTable() {
    // Это не нужно, т.к. дуструктор табицы дергает деструкторы дочерних элементов.
    // Да, Title - это дочерний. Если есть.
    /*
    if (Title) {
	    delete Title;
        Title = NULL;
    }
    */
}

String TExcelTable::GetTitle()
{
	if (!Title) return "";
	String out = Title->ReadValueString();
    return out;
}

TExcelTable* TExcelTable::SetTitle(const String& title) {
    if (!Title) return 0;
    Title->InsertString(title);
    return this;
}

TExcelCells* TExcelTable::GetHeader(unsigned int N)
{
	++N;

    vDataChild = vData.OlePropertyGet("HeaderRowRange").OlePropertyGet("Cells", 1, N); //.OlePropertyGet("Cells")
    //vDataChild.OleProcedure("Select");

    TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelCells* TExcelTable::GetHeaders() {
    vDataChild = vData.OlePropertyGet("HeaderRowRange");

    TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelTableColumn* TExcelTable::GetColumn(unsigned int N)
{
    vDataChild = vData.OlePropertyGet("ListColumns", N + 1);

    TExcelTableColumn* out = new TExcelTableColumn(this, vDataChild);
    return out;
}

unsigned int TExcelTable::ColumnsCount()
{
    return getChildCountByType("ListColumns");
}

TExcelCells* TExcelTable::GetField(unsigned int col, unsigned int row)
{
	col++;
	row++;
	vDataChild = vData.OlePropertyGet("ListRows", row).OlePropertyGet("Range").OlePropertyGet("Item", 1, col);
    TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

unsigned int TExcelTable::RowsCount()
{
	return getChildCountByType("ListRows");
}

TExcelTable* TExcelTable::AddRow()
{
    //if (pAddRow) pAddRow();
    vData.OlePropertyGet("ListRows").OleProcedure("Add");
    return this;
}

TExcelTable* TExcelTable::AddRows(TDataSet* src, const Variant& nullValue)
{
	unsigned int colCnt = src->FieldCount;
	unsigned int rowCnt = src->RecordCount;
	unsigned int startPos = getChildCountByType("ListRows");

	if (colCnt != ColumnsCount()) throw Exception("TExcelTable::AddRows: colCnt != ColumnsCount");

	String sNullVal = "";
	if (!nullValue.IsNull()) {
		sNullVal = VarToStr(nullValue);
	}

	unsigned int rPos = 1;
	for (src->First(); !src->Eof && rPos < rowCnt + 1; src->Next(), rPos++)
	{
		Variant vRow(OPENARRAY(int, (1, 1, 1, colCnt)), varVariant);
		Variant vElement;
		for (unsigned int j = 1; j < colCnt + 1; j++) {
			vElement = src->Fields->Fields[j - 1]->Value;

			CorrectInsert::InsertIntoVarArray(vElement, vRow, 1, j, sNullVal);
		}
		vData.OlePropertyGet("ListRows").OleProcedure("Add");
        vData.OlePropertyGet("ListRows", startPos + rPos).OlePropertyGet("Range").OlePropertySet("Value", vRow);
	}

    return this;
}

TExcelTable* TExcelTable::DeleteRow(unsigned int row)
{
	vData.OlePropertyGet("ListRows", row + 1).OleProcedure("Delete");
	return this;
}

}


