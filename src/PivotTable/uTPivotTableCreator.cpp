//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTPivotTableCreator.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
TPivotTableCreator::TPivotTableCreator(TExcelTable* source) 
{
    ChangeSourceTable(source);
}

TPivotTableCreator::~TPivotTableCreator() {
    SourceTable = NULL;
}

TExcelTable* TPivotTableCreator::initPivot(
    TExcelObjectRanged* Sheet, // ????
    unsigned int col, unsigned int row, 
    TPivotSettings* PivotSettings
    )
{
	Variant vPivotTable, vPivotField, vTitle;
    Variant vAim = Sheet->selectRange(col, row, col + PivotSettings->ColumnCount(), row + PivotSettings->RowCount())->getVariant();

	vPivotTable = Sheet->getParentVariant().OleFunction(
		"PivotTableWizard", // Функция, отвечающая за построение
		XlPivotTableSourceType::xlDatabase, // Тип источника (откуда)
		SourceTable->getVariant(), // Сам источник
		vAim, // Куда
		System::StringToOleStr(PivotSettings->NewSheetName), // Название
		PivotSettings->RowCount(),   // Кол-во строк
		PivotSettings->ColumnCount()   // Кол-во столбцов
	);
    
    for (unsigned int i = 0; i < PivotSettings->Settings.size(); ++i) {
		vPivotField = vPivotTable.OlePropertyGet("PivotFields", System::StringToOleStr(PivotSettings->Settings[i].ColumnName));

		vPivotField.OlePropertySet("Orientation", PivotSettings->Settings[i].Type);
		vPivotField.OlePropertySet("Caption", System::StringToOleStr(PivotSettings->Settings[i].Caption));
        
        if (PivotSettings->Settings[i].Type != XlPivotFieldOrientation::xlDataField) {
		    vPivotField.OlePropertySet("Position", 1);
        }
        else {
            vPivotField.OlePropertySet("Function", PivotSettings->Settings[i].Function);
        }
    }

	if (PivotSettings->DataPlace == PivotDataPlace::Columns) vPivotTable.OlePropertyGet("DataPivotField").OlePropertySet("Orientation", 2);

	//vCell = SelectCells(1, 1);


	if (PivotSettings->PivotTitle.Length() > 0){
        vTitle = Sheet->select(col, row)->getVariant();
		vTitle.OlePropertySet("Value", System::StringToOleStr(PivotSettings->PivotTitle));
	}

	//TExcelTable* table = new TExcelTable(Sheet, nullptr, nullptr, vPivotTable);
	return 0;
}

void TPivotTableCreator::ChangeSourceTable(TExcelTable *sourceNew)
{
    if (sourceNew == NULL) throw "TPivotTableCreator::ChangeSourceTable: SourceTable cannot be null!";
    SourceTable = sourceNew;
}

/*
TExcelTable* TPivotTableCreator::CreateTable(
    TExcelObjectRanged* sheet, 
    unsigned int col, unsigned int row, 
    TPivotSettings* PivotSettings
    )
{
    if (col < 1) col = 1;
	if (row < 3) row = 3;

    if (PivotSettings->needNewSheet){
		sheet = (TExcelObjectRanged*)sheet->getParent()->createChild();

        if (PivotSettings->NewSheetName.Length() < 1) {
            PivotSettings->NewSheetName = "Pivot_" + PivotSettings->GetPivotName();
        }

        sheet->SetName(PivotSettings->NewSheetName);
    }

    TExcelTable* out = initPivot(sheet, col, row, PivotSettings);
    return out;
}
*/

TExcelTable* TPivotTableCreator::CreateTable(
    unsigned int col, unsigned int row, 
    TPivotSettings* PivotSettings
    ) 
{
    if (col < 1) col = 1;
	if (row < 3) row = 3;

	TExcelObjectRanged* sheet = (TExcelObjectRanged*)SourceTable->getParent();
    /* if (PivotSettings->NewSheetName.Length() > 0){
        sheet = (TExcelObject*)(sheet->getParent())->Create();

        if (PivotSettings->NewSheetName.Length() < 1) {
            PivotSettings->NewSheetName = "Pivot_" + PivotSettings->GetPivotName();
        }

        sheet->SetName(PivotSettings->NewSheetName);
    } */

    return initPivot(sheet, col, row, PivotSettings);
}

}

