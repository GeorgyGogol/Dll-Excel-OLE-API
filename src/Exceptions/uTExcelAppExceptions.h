//---------------------------------------------------------------------------

#ifndef uTExcelAppExceptionsH
#define uTExcelAppExceptionsH

//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
class ExcelAppException : public ExcelException
{
public:
	ExcelAppException(const char* method, const char* message);
};

//---------------------------------------------------------------------------
class ExcelAppNotAttachedException : public ExcelAppException
{
public:
	ExcelAppNotAttachedException(const char* method);
};

//---------------------------------------------------------------------------
class ExcelAppAttachedException : public ExcelAppException
{
public:
	ExcelAppAttachedException(const char* method);
};

//---------------------------------------------------------------------------
#endif

