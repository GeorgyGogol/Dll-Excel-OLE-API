//---------------------------------------------------------------------------

#ifndef uTExcelSheetH
#define uTExcelSheetH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../PivotTable/uTPivotTableCreator.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelSheet : public TExcelObjectRanged {
public:
    TExcelSheet(TExcelObject* pParent, const Variant& data);
    TExcelSheet(const TExcelSheet&);
    ~TExcelSheet();

protected:
   	//void CreateTableOnCurrentCells(const String& tableName);	// Создать на текущих ячейках таблицу

public:
    // Одна ячейка
	TExcelCell* SelectCell(unsigned int col, unsigned int row);

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

	// Вставка с созданием НАСТОЯЩЕЙ ТАБЛИЦЫ как объекта
 	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet, const String& tableName, const String& tableTitle,
		bool needDisableSet = false
	);
 	
	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet, const String& tableName,
		bool needDisableSet = false
	);

	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDataSet* dataSet,
		bool needDisableSet = false
	);





//#ifdef DbgridehHPP
	TExcelTable* CreateTable(
		unsigned int startColumn, unsigned int startRow,
		TDBGridEh* gridEh, const String& tableName, const String& tableTitle,
		bool needDisableSet = false
	);
//#endif

	TExcelTable* CreatePivotTable(
		unsigned int startColumn, unsigned int startRow,
		TExcelTable* srcTable,
		TPivotSettings* pivotSettings
	);
	
	/* TExcelTable* CreatePivotTable(
		unsigned int startColumn, unsigned int startRow,
		const String& srcTable,
		const String& pivotName,
		TPivotSettings* pivotSettings
	); */


    // Место для диаграммм
};

//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif

