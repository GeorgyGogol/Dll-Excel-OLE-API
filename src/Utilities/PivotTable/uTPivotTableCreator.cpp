//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTPivotTableCreator.h"
#include "uTExcelSheet.h"

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
    TExcelObject* aimSheet,
    unsigned int col, unsigned int row, 
    TPivotSettings* PivotSettings
    )
{
	// TODO: Abomal program когда не находит поле
	Variant vPivotTable, vPivotField, vTitle;

	TExcelSheet* sheet = (TExcelSheet*)(aimSheet);

	if (PivotSettings->NewSheetName.Length() > 0) {
		//sheet = sheet->get
	}

	Variant vAim = sheet->SelectCells(col, row, col + PivotSettings->ColumnCount(), row + PivotSettings->RowCount())->getVariant();

	vPivotTable = sheet->GetParentVariant().OleFunction(
		"PivotTableWizard", // Функция, отвечающая за построение
		XlPivotTableSourceType::xlDatabase, // Тип источника (откуда)
		SourceTable->getVariant(), // Сам источник
		vAim, // Куда
		System::StringToOleStr(PivotSettings->GetPivotName()), // Название
		PivotSettings->ShowRowTotal,     // Строка итогов по строкам
		PivotSettings->ShowColumnTotal   // Строка итогов по столбцам
	);
    
	for (unsigned int i = 0; i < PivotSettings->Settings.size(); ++i) {
		TExcelTablePivotField cur = PivotSettings->Settings[i];
		vPivotField = vPivotTable.OlePropertyGet("PivotFields", System::StringToOleStr(cur.ColumnName));

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
		vTitle = sheet->SelectCell(1, 1)->getVariant();
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
    TExcelObject* aimSheet,
    unsigned int col, unsigned int row, 
    TPivotSettings* PivotSettings
    ) 
{
    if (col < 1) col = 1;
	if (row < 3) row = 3;

	//TExcelObjectRanged* sheet = (TExcelObjectRanged*)SourceTable->getParent();
    /* if (PivotSettings->NewSheetName.Length() > 0){
        sheet = (TExcelObject*)(sheet->getParent())->Create();

        if (PivotSettings->NewSheetName.Length() < 1) {
            PivotSettings->NewSheetName = "Pivot_" + PivotSettings->GetPivotName();
        }

        sheet->SetName(PivotSettings->NewSheetName);
    } */

	return initPivot(aimSheet, col, row, PivotSettings);
}

}

