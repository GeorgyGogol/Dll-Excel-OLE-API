//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObjectRanged.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
	
/* TExcelObjectRanged::TExcelObjectRanged()
	: TExcelObjectNode(), TExcelObjectData()
{} */

TExcelObjectRanged::TExcelObjectRanged(TExcelObject* pParent, const Variant& data)
	: TExcelObject(pParent, data)
{}

TExcelObjectRanged::TExcelObjectRanged(const TExcelObjectRanged& src)
	: TExcelObject(src)
{}

TExcelObjectRanged::~TExcelObjectRanged()
{}

AnsiString TExcelObjectRanged::ColToStrA(unsigned int ACol) {
	// Спизжено
	AnsiString out;

	// int x = (ACol-1) / 26;
	int k = (ACol - 1) / 26;
	int d = (ACol - 1) % 26;

	if (k)
		out.printf("%c", k + 64);
	out.cat_printf("%c", d + 65);

	return out;
}

AnsiString TExcelObjectRanged::GetRangeString(
	unsigned int startColumn, unsigned int startRow,
	unsigned int endColumn, unsigned int endRow
    )
{
	AnsiString out;

	if (endColumn == 0) endColumn = startColumn;
	if (endRow == 0) endRow = startRow;

	out.printf("%s%i:%s%i",	ColToStrA(startColumn).c_str(), 
							startRow, 
							ColToStrA(endColumn).c_str(), 
							endRow);

	return out;
}

AnsiString TExcelObjectRanged::GetCellString(unsigned int col, unsigned int row){
	AnsiString out;

	out.printf("%s%i", ColToStrA(col).c_str(), row);

	return out;
}

TExcelObjectRanged* TExcelObjectRanged::select(unsigned int col, unsigned int row) {
	seekAndSetDataChild("Range", GetCellString(col, row));
	vDataChild.OleProcedure("Select");

	return (TExcelObjectRanged*)getCurrentSelectedChild();
}

TExcelObjectRanged* TExcelObjectRanged::selectRange(
	unsigned int startColumn, unsigned int startRow,
	unsigned int endColumn, unsigned int endRow
	)
{
	AnsiString range = GetRangeString(startColumn, startRow, endColumn, endRow);
	seekAndSetDataChild("Range", range);
	vDataChild.OleProcedure("Select");

	return (TExcelObjectRanged*)getCurrentSelectedChild();
}

}

