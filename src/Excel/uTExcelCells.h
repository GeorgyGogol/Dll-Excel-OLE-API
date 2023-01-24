//---------------------------------------------------------------------------

#ifndef uTExcelCellsH
#define uTExcelCellsH

//---------------------------------------------------------------------------
#include "uTExcelObjectRangedTemplate.h"
//#include "uIFormatManager.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelCells : public TExcelObjectRangedTemplate<TExcelCells>//, public IFormatManager<TExcelCells>
{
public:
	TExcelCells(TExcelObject* pParent, const Variant& data);
    TExcelCells(TExcelCells&);
    ~TExcelCells();

private:
	unsigned int dColumn, dRow;

public:
    unsigned int GetColumnsCount();
    unsigned int GetRowCount();

	TExcelCells* GetCell(unsigned int col, unsigned int row);
	TExcelCells* GetCells(
	    unsigned int startColumn, unsigned int startRow,
	    unsigned int endColumn, unsigned int endRow
	);

    TExcelCells* Merge();

	TExcelCells* Insert(const Variant& data);
	TExcelCells* InsertString(const String& data);
	TExcelCells* InsertFormula(String& formula);

    Variant ReadValue();
    String ReadValueString();

	// IFormatManager
	TExcelCells* SetHorizontalAlign(ExcelTextAlign align);
	TExcelCells* SetVerticalAlign(ExcelTextAlign align);

	/*
	TExcelCells* SetBorders();

	TExcelCells* SetWidth();
	TExcelCells* SetHeight();
	TExcelCells* AutoSize();

	TExcelCells* SetFormat();
	*/
};

}
//---------------------------------------------------------------------------
#endif

