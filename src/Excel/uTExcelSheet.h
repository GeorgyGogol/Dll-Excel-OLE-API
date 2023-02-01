#ifndef uTExcelSheetH
#define uTExcelSheetH

#include "uTExcelCells.h"
#include "uTExcelTable.h"
#include "uTExcelNameItem.h"

#include "uTPivotTableCreator.h"
//---------------------------------------------------------------------------
namespace exl {
/** @addtogroup ExcelClientObjects
 * @{
 * @brief Класс Странички
 * 
 * Служит для взаимодействия со страницей. Предоставляет самые востребованные 
 * методы для работы с ними.
 */
class DLL_EI TExcelSheet : 
	public TExcelObjectRangedTemplate<TExcelSheet>,
	public ITExcelNames
{
public:
    TExcelSheet(TExcelObject* pParent, const Variant& data);
    TExcelSheet(const TExcelSheet&);
    ~TExcelSheet();

protected:
   	//void CreateTableOnCurrentCells(const String& tableName);	// Создать на текущих ячейках таблицу

public:
	TExcelCells* SelectCell(unsigned int col, unsigned int row); ///< Вернуть одну ячейку

    /// Вернуть несколько ячеек
	TExcelCells* SelectCells(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);

    // Для Колонок
	TExcelCells* SelectColumn(unsigned int column);
    //void HideColumn(unsigned int column);
    //void SetColumnWidth(unsigned int column, double size);

    // Для строк
	TExcelCells* SelectRow(unsigned int row);
    //void HideRow(unsigned int row);
    //void SetRowHeight(unsigned int row, double size);

	/// @brief Простая вставка таблицы (на самом деле это диапазон с данными)
	TExcelCells* InsertDataSet(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet, 
		bool needInsertFieldNames = false,
		bool needDisableSet = false
	);

	// Создать таблицу в диапазоне
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow);

	// Вставка с созданием НАСТОЯЩЕЙ ТАБЛИЦЫ как объекта в координаты
 	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet, const String& tableTitle, const String& tableName,
		bool needDisableSet = false
	);
 	
	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet, const String& tableTitle,
		bool needDisableSet = false
	);

	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet,
		bool needDisableSet = false
	);

	// Без координат
 	TExcelTable* CreateTable(
		TDataSet* dataSet, const String& tableTitle, const String& tableName,
		bool needDisableSet = false
	);
	
 	TExcelTable* CreateTable(
		TDataSet* dataSet, const String& tableTitle,
		bool needDisableSet = false
	);
 	
	TExcelTable* CreateTable(
		TDataSet* dataSet,
		bool needDisableSet = false
	);

	// Тут на основе GridEh
	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDBGridEh* gridEh, const String& tableTitle, const String& tableName,
		bool needDisableSet = false
	);
	
	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDBGridEh* gridEh, const String& tableTitle,
		bool needDisableSet = false
	);
	
	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDBGridEh* gridEh,
		bool needDisableSet = false
	);
	
	// Без координат
	TExcelTable* CreateTable(
		TDBGridEh* gridEh, const String& tableTitle, const String& tableName,
		bool needDisableSet = false
	);

	TExcelTable* CreateTable(
		TDBGridEh* gridEh, const String& tableTitle,
		bool needDisableSet = false
	);

	TExcelTable* CreateTable(
		TDBGridEh* gridEh,
		bool needDisableSet = false
	);

	// Сводная таблица
	TExcelTable* CreatePivotTable(
		unsigned int startColumn, unsigned int startRow,
		TExcelTable* srcTable,
		TPivotSettings* pivotSettings
	);

	TExcelTable* GetTable(const String& tableName);

	TExcelNameItem* GetNameItem(const String& itemName);
	TExcelNameItem* GetNameItem(unsigned int N);
	TExcelNameItem* AddNamedItem(const String& itemName);

    // Место для диаграммм
};

}
/// @}
#endif

