//---------------------------------------------------------------------------

#include <cmath>

#pragma hdrstop

#include "uTExcelObjectRangedTemplate.h"
#include "uTExcelDataExceptions.h"

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

/// @details Нагло украл у неизвестного программиста. В сути так и не разобрался
/// @param ACol Номер столбика
/// @return Буква столбика в мире Excel
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
		len = str.Length() - str.Pos(":") - secondPos + 1;
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

/// @details Проверяет. В зависимости от конфигурации, может сбрасывать исключение или
/// исправлять 0 на 1.
/// @param col Номер колонки
/// @param row Номер строчки
/// @warning Номерация начинается с 0
template<class T>
void TExcelObjectRangedTemplate<T>::checkColRow(unsigned int& col, unsigned int& row)
{
	if (col < 1) col = 1;
	if (row < 1) row = 1;
}

/// @param col Номер колонки
/// @param row Номер строчки
template<class T>
void TExcelObjectRangedTemplate<T>::selectSingle(unsigned int col, unsigned int row) {
	setSingle(col, row);
	vDataChild.OleProcedure("Select");
}

/// @brief 
/// @param startColumn Столбик, с которого выбираем
/// @param startRow Строка, с которой начинаем
/// @param endColumn Конечный столбик, по какой отбираем (включительно)
/// @param endRow Конечная строка, по которую отбираем (включительно)
/// @warning Номерация начинается с 0
/// @warning При конечных значениях меньше ячейки все равно отберутся. Ну, вроде должны
template<class T>
void TExcelObjectRangedTemplate<T>::selectRange(
	unsigned int startColumn, unsigned int startRow,
	unsigned int endColumn, unsigned int endRow
	)
{
	setRange(startColumn, startRow, endColumn, endRow);
	vDataChild.OleProcedure("Select");
}

template<class T>
void TExcelObjectRangedTemplate<T>::setSingle(unsigned int col, unsigned int row) {

	checkDataValide();
	checkColRow(col, row);
	seekAndSetDataChild("Range", GetCellString(col, row));
}

template<class T>
void TExcelObjectRangedTemplate<T>::setRange(
	unsigned int startColumn, unsigned int startRow,
	unsigned int endColumn, unsigned int endRow
	)
{
	checkDataValide();
	checkColRow(startColumn, startRow);
	checkColRow(endColumn, endRow);
	AnsiString range = GetRangeString(startColumn, startRow, endColumn, endRow);
	seekAndSetDataChild("Range", range);
}

// Для каждого - нужон шаблон
class DLL_EI TExcelSheet;
template class TExcelObjectRangedTemplate<TExcelSheet>;

class DLL_EI TExcelCells;
template class TExcelObjectRangedTemplate<TExcelCells>;

class DLL_EI TExcelTableColumn;
template class TExcelObjectRangedTemplate<TExcelTableColumn>;

class DLL_EI TExcelTable;
template class TExcelObjectRangedTemplate<TExcelTable>;

}


