//---------------------------------------------------------------------------

#ifndef uTExcelDataExceptionsH
#define uTExcelDataExceptionsH

//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
#define ERROR_EXCEL_DATA_ISNULL "Excel data is missing!"
#define ERROR_EXCEL_CELL_SELECT_DEFINED "Unable to select "
#define ERROR_EXCEL_CELL_SELECT_UNDEFINED "Unable select"
//---------------------------------------------------------------------------
class ExcelDataException : public ExcelException
{
public:
	ExcelDataException(const String& message)
		: ExcelException("Abstract", "ExcelData", message)
	{}

	ExcelDataException(const char* method, const char* message)
		: ExcelException("ExcelData", method, message)
	{}

	ExcelDataException(const char* message)
		: ExcelException("Abstract", "ExcelData", message)
	{}
};

//---------------------------------------------------------------------------
class ExcelDataNullException : public ExcelDataException
{
public:
	ExcelDataNullException()
		: ExcelDataException(ERROR_EXCEL_DATA_ISNULL)
	{}
};
//---------------------------------------------------------------------------
class ExcelSelectCellException : public ExcelDataException
{
public:
	ExcelSelectCellException()
		: ExcelDataException(ERROR_EXCEL_CELL_SELECT_UNDEFINED)
	{}

	ExcelSelectCellException(const char* type)
		: ExcelDataException(String(ERROR_EXCEL_CELL_SELECT_DEFINED) + String(type))
	{}
};

//---------------------------------------------------------------------------
#undef ERROR_EXCEL_DATA_ISNULL
#undef ERROR_EXCEL_CELL_SELECT
//---------------------------------------------------------------------------
#endif

