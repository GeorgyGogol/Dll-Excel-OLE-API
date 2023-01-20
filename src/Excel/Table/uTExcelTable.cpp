//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTable.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable)
    : TExcelObjectTemplate<TExcelTable>(pSheet, vTable)
{
    Variant vFields, vHeaders;
    vFields = vTable.OlePropertyGet("Range");
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

TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells)
	: TExcelObjectTemplate<TExcelTable>(pSheet, vTable)
		, Headers(headers), Data(data), Title(titleCells)
{
}

TExcelTable::TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells)
	: TExcelObjectTemplate<TExcelTable>(pSheet, Null())
		, Headers(headers), Data(data), Title(titleCells)
{
    //vData = 
}

TExcelTable::TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data)
	: TExcelObjectTemplate<TExcelTable>(pSheet, Null()), Headers(headers), Data(data), Title(NULL)
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
    //Parent->RemoveChildClass(this);
	delete Title;
    delete Headers;
    delete Data;
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

TExcelCells* TExcelTable::GetHeader(unsigned int N) {
    vDataChild = vData.OlePropertyGet("HeaderRowRange").OlePropertyGet("Cells", 1, N); //.OlePropertyGet("Cells")
    //vDataChild.OleProcedure("Select");

    TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

TExcelTableColumn* TExcelTable::GetColumn(unsigned int N)
{
    vDataChild = vData.OlePropertyGet("ListColumns", N);

    TExcelTableColumn* out = new TExcelTableColumn(this, vDataChild);
    return out;
}

TExcelCells* TExcelTable::GetField(unsigned int col, unsigned int row)
{
    vDataChild = vData.OlePropertyGet("ListRows", row).OlePropertyGet("Range").OlePropertyGet("Item", 1, col);
    TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

void* TExcelTable::AddRow()
{
    vDataChild = vData.OlePropertyGet("ListRows").OleProcedure("Add");
    //return 
}

//---------------------------------------------------------------------------
}

