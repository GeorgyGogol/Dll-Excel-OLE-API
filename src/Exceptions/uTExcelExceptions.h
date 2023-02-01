//---------------------------------------------------------------------------

#ifndef uTExcelExceptionsH
#define uTExcelExceptionsH

//---------------------------------------------------------------------------
#include "uDll.h"
//---------------------------------------------------------------------------
class ExcelException : public Exception
{
public:
	ExcelException(const String& message);
	ExcelException(const char* message);

protected:
	ExcelException(const String& classWhere, const String& method, const String& message);
	ExcelException(const char* classWhere, const char* method, const char* message);
};

//---------------------------------------------------------------------------
#endif

