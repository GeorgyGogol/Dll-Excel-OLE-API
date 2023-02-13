//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTable.h"
#include "uFunctions.h"
#include "uCheckers.h"

#include "Log.h"
#include "uBuildSettings.h"

#define MBYTES_TO_BYTES(mb) mb * 1024 * 1024

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable)
	: TExcelObjectRangedTemplate<TExcelTable>(pSheet, vTable)
{
}

TExcelTable::TExcelTable(TExcelObject* pSheet, const Variant& vTable, TExcelCells* eTitle)
	: TExcelObjectRangedTemplate<TExcelTable>(pSheet, vTable)
{
	if (eTitle) {
		Title = new TExcelCells(this, eTitle->getVariant());
	}
}

TExcelTable::~TExcelTable() 
{
    /// Так как Title - дочерний элемент, удалится сам.
}

const unsigned int TExcelTable::decideOneStepRows(const unsigned int& nCols) const
{
	// Кол-во памяти / размер объекта (= кол-во записей) / кол-во колонок
	return MBYTES_TO_BYTES(32) / sizeof(Variant) / nCols;
	// Кол-во записей / кол-во колонок
	//return 500000 / src->FieldCount; /* (500к ~ 351 КБ) */
}

/// @return Содержимое ячейки (объединенных ячеек)
String TExcelTable::GetTitle()
{
	if (!Title) return "";
	String out = Title->ReadValueString();
    return out;
}

/// @details Установить заголовок таблицы, если есть
/// @param title Содержимое заголовка
TExcelTable* TExcelTable::SetTitle(const String& title) {
    if (!Title) return 0;
    Title->InsertString(title);
    return this;
}

/// @param N Номер заголовка от 0 до N-1 включительно
/// @note Отсчет идет по-программистки, с 0, а не по-людски.
/// @return Ячейка заголовка
TExcelCells* TExcelTable::GetHeader(unsigned int N)
{
    /// Обратится к заголовку колонки - свойству HeaderRowRange, свойству Cells, ячейке (1, N)
    vDataChild = vData.OlePropertyGet("HeaderRowRange").OlePropertyGet("Cells", 1, N + 1);
    TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

/// @return Ячейки заголовков таблиц
TExcelCells* TExcelTable::GetHeaders() {
    vDataChild = vData.OlePropertyGet("HeaderRowRange");

    TExcelCells* out = new TExcelCells(this, vDataChild);
	return out;
}

/// @param N Номер столбика-колонки
/// @return Экземпляр класса колонки со всеми ячейками (заголовки не включены)
TExcelTableColumn* TExcelTable::GetColumn(unsigned int N)
{
    /// Для этого идет обращение к свойству ListColumns
    vDataChild = vData.OlePropertyGet("ListColumns", N + 1);
    TExcelTableColumn* out = new TExcelTableColumn(this, vDataChild);
    return out;
}

unsigned int TExcelTable::ColumnsCount()
{
    return getChildCountByType("ListColumns");
}

/// @param col Столбец/Колонка
/// @param row Номер записи/строчик
/// @return Ячейку
TExcelCells* TExcelTable::GetField(unsigned int col, unsigned int row)
{
    /// @note Отсчет идет по-программистки, с 0, а не по-людски. Есть
    /// поправка на +1.
	col++;
	row++;

    /// Для этого обратимся к свойству ListRows, получим ListRow. Обратимся к его
    /// свойству Range, потом уже свойству Item (1, col). 1 - потому что это количество
    /// строк, на сколько понял.
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
	/// У свойства ListRows дергается процедура (метод) Add, который добавляет
    /// строку снизу таблицы.
	/// @note При добавлении строки добавляются новые строчки в таблицу, лист
	/// "расширяется" вниз
	vDataChild = vData.OlePropertyGet("ListRows").OleFunction("Add").OlePropertyGet("Range");
	return this;
}

TExcelTable* TExcelTable::AddRow(unsigned int pos)
{
	checkers::checkNumber(pos);
	vDataChild = vData.OlePropertyGet("ListRows").OleFunction("Add", pos).OlePropertyGet("Range");
	return this;
}

TExcelTable* TExcelTable::AddRows(unsigned int from, unsigned int cnt)
{
	checkers::checkNumber(from);
	if (!checkers::checkNumber(cnt)) throw "Cannot add less than 1 row";

	String StartRowString = vData.OlePropertyGet("ListRows", from).OlePropertyGet("Range").OlePropertyGet("Address");

	unsigned int StartRow = GetRowFromStr(StartRowString) + (from - 1);
	unsigned int StartColumn = GetColFromStr(StartRowString);
	String end = StartRowString.SubString(StartRowString.Pos(":") + 1, StartRowString.Length() - StartRowString.Pos(":"));
	unsigned int endColumn = GetColFromStr(end);

	String sRange = GetRangeString(StartColumn, StartRow, endColumn, StartRow + cnt - 1);

	vDataChild = GetParentVariant().OlePropertyGet("Range", System::StringToOleStr(sRange));
	vDataChild.OleProcedure("Insert", -4121);
	vDataChild = GetParentVariant().OlePropertyGet("Range", System::StringToOleStr(sRange));
	return this;
}

TExcelTable* TExcelTable::AddRows(unsigned int cnt)
{
	int Last = getChildCountByType("ListRows");
	return AddRows(Last, cnt);
}

TExcelTable* TExcelTable::AddRows(TDataSet* src, const Variant& nullValue)
{
	TLog log("TExcelTable::AddRows");
	const unsigned int RealSourceRowCount = src->RecordCount;
	unsigned int colCnt = src->FieldCount;

	if (colCnt < getChildCountByType("ListColumns")) {
		log.WriteLog("Table have more Columns, what Source. Extra Columns will be truncated", CallLogReason::Warning);
	}
	else if (colCnt > getChildCountByType("ListColumns")) {
		colCnt = getChildCountByType("ListColumns");
		log.WriteLog("Table have less Columns, what Source. Extra Columns will be truncated", CallLogReason::Warning);
	}

	const unsigned int OneStepReadRows = decideOneStepRows(colCnt);
	unsigned int nSteps = static_cast<unsigned int>(RealSourceRowCount / OneStepReadRows);

	if (nSteps * OneStepReadRows < RealSourceRowCount) {
		++nSteps;
	}

	String sNullVal = "";
	if (!nullValue.IsNull()) {
		sNullVal = VarToStr(nullValue);
	}

	#ifdef ENABLE_USAGE_STATISTIC
	log.WriteLog("Run");
	log.WriteLog("Iterations", nSteps);
	#endif

	unsigned int startPos, endPos;
	src->First();
	for (unsigned int step = 0; step < nSteps; ++step) {
		
		#ifdef ENABLE_USAGE_STATISTIC
		log.WriteLog("Iteration", step);
		#endif

	   	startPos = step * OneStepReadRows;
		endPos = (step + 1) * OneStepReadRows;

		if (endPos > RealSourceRowCount) endPos = RealSourceRowCount;
		--endPos;

		Variant varArray(OPENARRAY(int, (1, endPos - startPos + 1, 1, colCnt)), varVariant);

		for (unsigned int rPos = 1; !src->Eof && rPos < endPos - startPos + 1 + 1; src->Next(), ++rPos) {
			for (unsigned int tabCol = 0; tabCol < colCnt; ++tabCol) {
				CorrectInsert::InsertIntoVarArray(src->Fields->Fields[tabCol]->Value, varArray, rPos, tabCol + 1, sNullVal);
			}
		}

		AddRows(endPos - startPos + 1);
		vDataChild.OlePropertySet("Value", varArray);

		#ifdef ENABLE_USAGE_STATISTIC
		log.WriteSize(sizeof(Variant) * colCnt * (endPos - startPos) + sizeof(varArray));
		log.WriteLog("Iteration done");
		#endif
	}

	#ifdef ENABLE_USAGE_STATISTIC
	log.WriteLog("End");
	#endif

    return this;
}

TExcelTable* TExcelTable::DeleteRow(unsigned int row)
{
    /// @note Отсчет идет по-программистки, с 0, а не по-людски.
    /// @note При удалении строки лист "подтягивается" вверх
	vData.OlePropertyGet("ListRows", row + 1).OleProcedure("Delete");
	return this;
}

}


