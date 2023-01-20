//---------------------------------------------------------------------------

#ifndef uTExcelTableH
#define uTExcelTableH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelTableHeaders.h"
//---------------------------------------------------------------------------
namespace exl {

class TExcelTableColumn : public TExcelObjectRangedTemplate<TExcelTableColumn>
{
public:
	TExcelTableColumn(TExcelObject* pParent, const Variant& data)
		: TExcelObjectRangedTemplate<TExcelTableColumn>(pParent, data)
		{}

    TExcelTableColumn(TExcelTableColumn& src)
		: TExcelObjectRangedTemplate<TExcelTableColumn>(src)
	{}

    ~TExcelTableColumn() {}

public:
	TExcelTableColumn* SetIdentity(int start, int step){
		vData.OlePropertyGet("Range").OlePropertySet("FormulaR1C1", "=R[-1]C+1");
		vData.OlePropertyGet("Range").OlePropertyGet("Item", 2).OlePropertySet("FormulaR1C1", "=1");
	}

	TExcelTableColumn* SetBorders();
	TExcelTableColumn* SetWidth();
	TExcelTableColumn* SetFormat();
};

//---------------------------------------------------------------------------
class DLL_EI TExcelTable : public TExcelObjectTemplate<TExcelTable>
{
public:
	TExcelTable(TExcelObject* pSheet, const Variant& vTable);

	TExcelTable(TExcelObject* pSheet, const Variant& vTable, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells);

	//TExcelTable(TExcelObjectRanged* sheet, const Variant& name, const Variant& headers, const Variant& data);
	TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells);
	TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data);
    //TExcelTable(TExcelObjectRanged* sheet, TExcelCell* cells, unsigned int headersCount);
    ~TExcelTable();

private:
	TExcelCells* Title;
	TExcelTableHeaders* Headers;
	TExcelCells* Data;

public:
	String GetTitle();
    TExcelTable* SetTitle(const String& title);

	// Headers
	TExcelCells* GetHeader(unsigned int N);

	// Columns
	TExcelTableColumn* GetColumn(unsigned int N);

	// Records and fields
	TExcelCells* GetField(unsigned int col, unsigned int row);

	void* AddRow();

	// ListObject.ShowTotals property

    //TExcelCell* GetCell(unsigned int col, unsigned int row);
};

//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif


