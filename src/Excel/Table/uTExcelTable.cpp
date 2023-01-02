//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTable.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
/*
TExcelTable::TExcelTable(TExcelObjectRanged* Sheet, const Variant& name, const Variant& headers, const Variant& data)
	: TExcelObjectRanged(*Sheet)
{
	NameCells = new TExcelCell(Sheet, name);
	Headers = new TExcelTableHeaders(Sheet, headers);
	Data = new TExcelTableData(Sheet, data);
}
*/

TExcelTable::TExcelTable(TExcelObjectRanged* Sheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCell* titleCells)
	: TExcelObject(Sheet->getParent(), Null())
        , Headers(headers), Data(data), TitleCells(titleCells)
{
}

TExcelTable::TExcelTable(TExcelObjectRanged* Sheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data)
	: TExcelObject(Sheet->getParent(), Null()), Headers(headers), Data(data), TitleCells(NULL)
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
    delete TitleCells;
    delete Headers;
    delete Data;
}

String TExcelTable::GetTitle()
{
    if (!TitleCells) return "";
	String out = TitleCells->ReadString();
    return out;
}

TExcelTable* TExcelTable::SetTitle(const String& title) {
    if (!TitleCells) return 0;
    TitleCells->InsertString(title);
    return this;
}

TExcelTableHeaders* TExcelTable::GetHeaders() {
    return Headers;
}

//---------------------------------------------------------------------------
}

