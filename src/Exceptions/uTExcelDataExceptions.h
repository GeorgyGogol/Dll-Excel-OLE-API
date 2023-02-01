//---------------------------------------------------------------------------

#ifndef uTExcelDataExceptionsH
#define uTExcelDataExceptionsH

//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
class ExcelDataException : public ExcelException
{
public:
	ExcelDataException(const String& message);
	ExcelDataException(const char* method, const char* message);
	ExcelDataException(const char* message);
};

//---------------------------------------------------------------------------
class ExcelDataNullException : public ExcelDataException
{
public:
	ExcelDataNullException();
};
//---------------------------------------------------------------------------
class ExcelSelectCellException : public ExcelDataException
{
public:
	ExcelSelectCellException();
	ExcelSelectCellException(const char* type);
};

//---------------------------------------------------------------------------
#endif

