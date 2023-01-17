//---------------------------------------------------------------------------

#ifndef ExcelAPIH
#define ExcelAPIH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
// Settings:
//#define EXCEL_APP_CREATED_ERROR
#define EXCEL_SAVE_CELLS_SELECT
//---------------------------------------------------------------------------

//#include "uDll.h"
//#include "Excel/uTExcelEnums.h"
//#include "Excel/Abstract/SubClasses/uTExcelObjectNode.h"
//#include "Excel/Abstract/SubClasses/uTExcelObjectData.h"
//#include "Excel/Abstract/uTExcelObject.h"
//#include "Excel/Abstract/uTExcelObjectTemplate.h"
//#include "Excel/Abstract/uTExcelObjectRangedTemplate.h"
//#include "Excel/uTExcelCells.h"
//#include "Excel/Table/uTExcelTableHeaders.h"
//#include "Excel/Table/uTExcelTable.h"
//#include "Utilities/Table/uTGridHeaders.h"
//#include "Utilities/Table/uTTableCreator.h"
//#include "Utilities/PivotTable/uTPivotSettings.h"
//#include "Utilities/PivotTable/uTPivotTableCreator.h"
//#include "Excel/uTExcelSheet.h"
//#include "Excel/uTExcelWorkbook.h"

#include "Excel/uTExcelApp.h"

extern "C" {
namespace exl {

class DLL_EI TExcelObject;
class DLL_EI TExcelCell;
class DLL_EI TExcelCells;
//class DLL_EI TExcelTableHeaders;
class DLL_EI TExcelTable;
//class DLL_EI TTableCreator;
//struct DLL_EI TExcelTablePivotField;
//class DLL_EI TPivotSettings;
//class DLL_EI TPivotTableCreator;
class DLL_EI TExcelSheet;
class DLL_EI TExcelWorkbook;
class DLL_EI TExcelApp;

}
}

//---------------------------------------------------------------------------
#endif

