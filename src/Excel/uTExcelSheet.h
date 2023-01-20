//---------------------------------------------------------------------------

#ifndef uTExcelSheetH
#define uTExcelSheetH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../Utilities/PivotTable/uTPivotTableCreator.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelSheet : public TExcelObjectRangedTemplate<TExcelSheet> 
{
public:
    TExcelSheet(TExcelObject* pParent, const Variant& data);
    TExcelSheet(const TExcelSheet&);
    ~TExcelSheet();

protected:
   	//void CreateTableOnCurrentCells(const String& tableName);	// Создать на текущих ячейках таблицу

public:
    // Одна ячейка
	TExcelCells* SelectCell(unsigned int col, unsigned int row);

    // Несколько ячеек
	TExcelCells* SelectCells(
		unsigned int startColumn, unsigned int startRow,
		unsigned int endColumn, unsigned int endRow
    );

    // Для Колонок
	TExcelCells* SelectColumn(unsigned int column);
    //void HideColumn(unsigned int column);
    //void SetColumnWidth(unsigned int column, double size);

    // Для строк
	TExcelCells* SelectRow(unsigned int row);
    //void HideRow(unsigned int row);
    //void SetRowHeight(unsigned int row, double size);

	// Простая вставка таблицы (на самом деле это диапазон с данными)
	TExcelCells* InsertDataSet(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet, 
		bool needInsertFieldNames = false,
		bool needDisableSet = false
	);

	// Создать таблицу в диапозоне
	// TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow);

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

	// Найти и вернуть таблицу
	TExcelTable* GetTable(const String& tableName);

    // Место для диаграммм
};

//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif

