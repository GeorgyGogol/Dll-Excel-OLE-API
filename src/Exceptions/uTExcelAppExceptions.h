//---------------------------------------------------------------------------

#ifndef uTExcelAppExceptionsH
#define uTExcelAppExceptionsH

//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
#define ERROR_EXCEL_NOT_CREATED "Excel app is not Created or Attached!"
#define ERROR_EXCEL_CREATED "Cannot create or attach one more Excel in this object! Please, use one TExcelApp to one ExcelInstance"
//---------------------------------------------------------------------------
class ExcelAppException : public ExcelException
{
public:
	ExcelAppException(const char* method, const char* message)
		: ExcelException("TExcelApp", method, message)
	{}
};
//---------------------------------------------------------------------------
class ExcelAppNotAttachedException : public ExcelAppException
{
public:
	ExcelAppNotAttachedException(const char* method)
		: ExcelAppException(method, ERROR_EXCEL_NOT_CREATED)
	{}
};

//---------------------------------------------------------------------------
class ExcelAppAttachedException : public ExcelAppException
{
public:
	ExcelAppAttachedException(const char* method)
		: ExcelAppException(method, ERROR_EXCEL_CREATED)
	{}
};

//---------------------------------------------------------------------------
#undef ERROR_EXCEL_NOT_CREATED
#undef ERROR_EXCEL_CREATED
//---------------------------------------------------------------------------
#endif

