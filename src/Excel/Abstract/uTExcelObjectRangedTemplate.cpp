//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObjectRangedTemplate.h"
#include "../../Exceptions/uTExcelDataExceptions.h"
#include <cmath>

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
template<class T>
TExcelObjectRangedTemplate<T>::TExcelObjectRangedTemplate(TExcelObject* pParent, const Variant& data)
	: TExcelObjectTemplate<T>(pParent, data)
{}

template<class T>
TExcelObjectRangedTemplate<T>::TExcelObjectRangedTemplate(const TExcelObjectRangedTemplate<T>& src)
	: TExcelObjectTemplate<T>(src)
{}

template<class T>
TExcelObjectRangedTemplate<T>::~TExcelObjectRangedTemplate()
{}

template<class T>
AnsiString TExcelObjectRangedTemplate<T>::ColToStrA(unsigned int ACol) {
	// Нагло спизжено
	AnsiString out;

	// int x = (ACol-1) / 26;
	int k = (ACol - 1) / 26;
	int d = (ACol - 1) % 26;

	if (k)
		out.printf("%c", k + 64);
	out.cat_printf("%c", d + 65);

	return out;
}

template<class T>
int TExcelObjectRangedTemplate<T>::ColStrToInteger(const String& str) const
{
	// Нагло придумано
	if (str.Length() > 1) {
		unsigned int add = ((int)(str.SubString(1, 1).c_str()[0])) - 64;
		unsigned int multi = std::pow(26.0, str.Length() - 1);
		return ColStrToInteger(str.SubString(2, str.Length() - 1)) + (add * multi);
	}
	else {
		return ((int)(str.c_str()[0])) - 64;
	}
}

template<class T>
unsigned int TExcelObjectRangedTemplate<T>::GetColFromStr(const String& str) {
	unsigned int out = 0;

	String aim = str.SubString(2, str.SubString(2, str.Length() - 1).Pos("$") - 1);

	out = ColStrToInteger(aim);

	return out;
}

template<class T>
unsigned int TExcelObjectRangedTemplate<T>::GetRowFromStr(const String& str) {
	int secondPos = str.SubString(2, str.Length()).Pos("$") + 2;
	int len = str.Length() - secondPos;

	if (str.Pos(":")) {
		len -= (str.Pos(":") + 1);
	}

	String aim = str.SubString(secondPos, len);
	return (unsigned int)StrToInt(aim);
}

template<class T>
AnsiString TExcelObjectRangedTemplate<T>::GetRangeString(
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

template<class T>
AnsiString TExcelObjectRangedTemplate<T>::GetCellString(unsigned int col, unsigned int row){
	AnsiString out;

	out.printf("%s%i", ColToStrA(col).c_str(), row);

	return out;
}

template<class T>
void TExcelObjectRangedTemplate<T>::checkColRow(unsigned int& col, unsigned int& row)
{
#ifdef EXCEL_SAVE_CELLS_SELECT
	if (col < 1) col = 1;
	if (row < 1) row = 1;
#else
	if (col < 1) throw ExcelDataException("Requested Column is less that 1");
	if (row < 1) throw ExcelDataException("Requested Row is less than 1");
#endif
}

template<class T>
void TExcelObjectRangedTemplate<T>::selectSingle(unsigned int col, unsigned int row) {
	checkDataValide();
	checkColRow(col, row);
	seekAndSetDataChild("Range", GetCellString(col, row));
	vDataChild.OleProcedure("Select");
}

template<class T>
void TExcelObjectRangedTemplate<T>::selectRange(
	unsigned int startColumn, unsigned int startRow,
	unsigned int endColumn, unsigned int endRow
	)
{
	checkDataValide();
	checkColRow(startColumn, startRow);
	checkColRow(endColumn, endRow);
	AnsiString range = GetRangeString(startColumn, startRow, endColumn, endRow);
	seekAndSetDataChild("Range", range);
	vDataChild.OleProcedure("Select");
}

// Для каждого - нужон шаблон
class DLL_EI TExcelSheet;
template class TExcelObjectRangedTemplate<TExcelSheet>;

class DLL_EI TExcelCells;
template class TExcelObjectRangedTemplate<TExcelCells>;


}

