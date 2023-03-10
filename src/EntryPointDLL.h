#ifndef EntryPointDLLH
#define EntryPointDLLH

//---------------------------------------------------------------------------

#include "uDll.h"

// interfaces
#include "uIFormatManager.h"
#include "uICreateTable.h"
#include "uIGetTable.h"
#include "uIBorderManager.h"


extern "C" {

#include "Excel/Abstract/SubClasses/uTExcelObjectNode.h"
#include "Excel/Abstract/SubClasses/uTExcelObjectData.h"
#include "Excel/Abstract/uTExcelObject.h"
}

#include "Excel/Abstract/uTExcelObjectTemplate.h"
#include "Excel/Abstract/uTExcelObjectRangedTemplate.h"

extern "C" {

#include "Excel/uTExcelCells.h"
#include "Excel/Table/uTExcelTable.h"
#include "Utilities/Table/uTGridHeaders.h"
#include "Utilities/Table/uTTableCreator.h"
#include "Utilities/PivotTable/uTPivotSettings.h"
#include "Utilities/PivotTable/uTPivotTableCreator.h"
#include "Excel/uTExcelSheet.h"
#include "Excel/uTExcelWorkbook.h"
#include "Excel/uTExcelApp.h"

}

//---------------------------------------------------------------------------
#endif
