//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelDataExceptions.h"

//---------------------------------------------------------------------------
#define ERROR_EXCEL_DATA_ISNULL "Excel data is missing!"
#define ERROR_EXCEL_CELL_SELECT_DEFINED "Unable to select "
#define ERROR_EXCEL_CELL_SELECT_UNDEFINED "Unable select"
//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
ExcelDataException::ExcelDataException(const String& message)
    : ExcelException("Abstract", "ExcelData", message)
{}

ExcelDataException::ExcelDataException(const char* method, const char* message)
    : ExcelException("ExcelData", method, message)
{}

ExcelDataException::ExcelDataException(const char* message)
    : ExcelException("Abstract", "ExcelData", message)
{}

ExcelDataNullException::ExcelDataNullException()
    : ExcelDataException(ERROR_EXCEL_DATA_ISNULL)
{}

/* ExcelSelectCellException::ExcelSelectCellException()
    : ExcelDataException(ERROR_EXCEL_CELL_SELECT_UNDEFINED)
{}

ExcelSelectCellException::ExcelSelectCellException(const char* type)
    : ExcelDataException(String(ERROR_EXCEL_CELL_SELECT_DEFINED) + String(type))
{} */

