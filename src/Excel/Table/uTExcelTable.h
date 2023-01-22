//---------------------------------------------------------------------------

#ifndef uTExcelTableH
#define uTExcelTableH

//---------------------------------------------------------------------------
#include "uDll.h"
#include "uTExcelCells.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelTableColumn : public TExcelObjectRangedTemplate<TExcelTableColumn>
                               //, public IFormatManager<TExcelTableColumn>
{
public:
	TExcelTableColumn(TExcelObject* pParent, const Variant& data);
    TExcelTableColumn(TExcelTableColumn& src);
    ~TExcelTableColumn();

public:
	TExcelTableColumn* SetIdentity(int start, int step);

	TExcelTableColumn* SetHorizontalAlign(ExcelTextAlign align);
	TExcelTableColumn* SetVerticalAlign(ExcelTextAlign align);

	TExcelTableColumn* SetBorders();

	TExcelTableColumn* SetWidth();
	TExcelTableColumn* SetHeight();
	TExcelTableColumn* AutoSize();

	TExcelTableColumn* SetFormat();

};

//---------------------------------------------------------------------------
class DLL_EI TExcelTable : public TExcelObjectTemplate<TExcelTable>
{
public:
	TExcelTable(TExcelObject* pSheet, const String& tableName);
	TExcelTable(TExcelObject* pSheet, const Variant& vTable);
	//TExcelTable(TExcelObject* pSheet, TExcelCells* eCells);
	TExcelTable(TExcelObject* pSheet, const Variant& vTable, const Variant& vTitle);
	TExcelTable(TExcelObject* pSheet, const Variant& vTable, TExcelCells* eTitile);
	//TExcelTable(TExcelObject* pSheet, const String& tableName, const Variant& vTitle);


	//TExcelTable(TExcelObject* pSheet, const Variant& vTable, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells);

	//TExcelTable(TExcelObjectRanged* sheet, const Variant& name, const Variant& headers, const Variant& data);
	//TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCells* titleCells);
	//TExcelTable(TExcelObject* pSheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data);
	//TExcelTable(TExcelObjectRanged* sheet, TExcelCell* cells, unsigned int headersCount);

	~TExcelTable();

private:

private:
	TExcelCells* Title;

public:
	String GetTitle();
    TExcelTable* SetTitle(const String& title);

	// Headers
	TExcelCells* GetHeader(unsigned int N);
	TExcelCells* GetHeaders();

	// Columns
	TExcelTableColumn* GetColumn(unsigned int N);

	unsigned int ColumnsCount();

	// Records and fields
	TExcelCells* GetField(unsigned int col, unsigned int row);

	unsigned int RowsCount();

	TExcelTable* AddRow();
	TExcelTable* AddRows(TDataSet* src, const Variant& nullValue = Null());
	TExcelTable* DeleteRow(unsigned int row);

	// ListObject.ShowTotals property

    //TExcelCell* GetCell(unsigned int col, unsigned int row);
};

}
//---------------------------------------------------------------------------
#endif

