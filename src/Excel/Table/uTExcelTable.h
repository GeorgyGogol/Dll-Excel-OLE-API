#ifndef uTExcelTableH
#define uTExcelTableH

#include "uTExcelTableColumn.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс таблицы
 * @addtogroup Tables
 * 
 * Предназначен для удобного управления объектом таблицы.
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelTable : public TExcelObjectRangedTemplate<TExcelTable>
{
public:
	TExcelTable(TExcelObject* pSheet, const Variant& vTable); ///< Создать отражение объекта с этой вариантой
	TExcelTable(TExcelObject* pSheet, const Variant& vTable, TExcelCells* eTitile); ///< Создать отражение объекта с этой вариантой и заголовком
	~TExcelTable();

private:
	TExcelCells* Title; ///< Указатель на ячейки с заголовком

	const unsigned int decideOneStepRows(	const unsigned int& nColumns) const;

public:
	String GetTitle(); ///< Получить текст заголовка
    TExcelTable* SetTitle(const String& title); ///< Установить текст заголовка

	// Headers
	TExcelCells* GetHeader(unsigned int N); ///< Обратиться к определенному заголовку столбика
	TExcelCells* GetHeaders(); ///< Получить весь диапозон заголовков столбцов

	// Columns
	TExcelTableColumn* GetColumn(unsigned int N); ///< Обратиться к определенной колонке
	unsigned int ColumnsCount(); ///< Получить кол-во колонок

	// Records and fields
	TExcelCells* GetField(unsigned int col, unsigned int row); ///< Получить определенное поле таблицы

	unsigned int RowsCount(); ///< Получить кол-во строк
 
	TExcelTable* AddRow(); ///< Добавить запись в конец
	TExcelTable* AddRow(unsigned int pos); ///< Добавить запись на позицию
	TExcelTable* AddRows(unsigned int cnt); ///< Добавить несколько записей в конец
	TExcelTable* AddRows(unsigned int from, unsigned int cnt); ///< Добавить записи на позицию

	TExcelTable* AddRows(TDataSet* src, const Variant& nullValue = Null()); ///< Добавить записи в конец из TDataSet
	
	TExcelTable* DeleteRow(unsigned int row); ///< Удалить запись (строчку)

};

}

//---------------------------------------------------------------------------
#endif

