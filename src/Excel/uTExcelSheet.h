#ifndef uTExcelSheetH
#define uTExcelSheetH

#include "uTExcelCells.h"
#include "uACreateTableController.h"
#include "uTExcelNameItem.h"
#include "uIGetTable.h"

#include "uTPivotTableCreator.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс Странички
 * @ingroup ExcelClientObjects
 * 
 * Служит для взаимодействия со страницей. Предоставляет самые востребованные 
 * методы для работы с ними.
 * 
 * Методы Select - делают дополнительное действие - выбрать
 * Методы Get - просто возвращают ячейки
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelSheet : 
	public TExcelObjectRangedTemplate<TExcelSheet>,
	public ACreateTableController,
	public ITExcelNames,
	public IGetTable<TExcelTable>
{
public:
    TExcelSheet(TExcelObject* pParent, const Variant& data);
    TExcelSheet(const TExcelSheet&);
    ~TExcelSheet();

public:

	TExcelCells* GetCell(unsigned int col, unsigned int row);
	TExcelCells* GetCell(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);


	TExcelCells* SelectCell(unsigned int col, unsigned int row);
	TExcelCells* SelectCells(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);

	/// @brief Выбрать и получить столбец
	/// @return Выбранный столбик
	TExcelCells* SelectColumn(unsigned int column);
	TExcelCells* SelectColumns(unsigned int colStart, unsigned int colEnd);

	/// @brief Выбрать и получить строку
	/// @return Выбранная строка
	TExcelCells* SelectRow(unsigned int row);
	TExcelCells* SelectRows(unsigned int rowStart, unsigned int rowEnd);

	/// @brief Простая вставка таблицы (на самом деле это диапазон с данными)
	TExcelCells* InsertDataSet(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet, bool needInsertFieldNames = false);

	// ACreateTableController override
	// Вставка с созданием НАСТОЯЩЕЙ ТАБЛИЦЫ как объекта в координаты
 	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet, const String& tableTitle, const String& tableName);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet, const String& tableTitle);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet);
	// Тут на основе GridEh
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDBGridEh* gridEh, const String& tableTitle, const String& tableName);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDBGridEh* gridEh, const String& tableTitle);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDBGridEh* gridEh);

	// IGetTable override
	TExcelTable* GetTable(const String& tableName);


	/// Сводная таблица
	TExcelTable* CreatePivotTable(
		unsigned int startColumn, unsigned int startRow,
		TExcelTable* srcTable,
		TPivotSettings* pivotSettings
	);

	// ITExcelNames override
	TExcelNameItem* GetNameItem(const String& itemName);
	TExcelNameItem* GetNameItem(unsigned int N);
	TExcelNameItem* AddNamedItem(const String& itemName);

};

}

//---------------------------------------------------------------------------
#endif

