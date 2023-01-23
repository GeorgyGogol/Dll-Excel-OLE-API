//---------------------------------------------------------------------------

#ifndef uTExcelExceptionsH
#define uTExcelExceptionsH

//---------------------------------------------------------------------------
#include "uDll.h"
//---------------------------------------------------------------------------
class ExcelException : public Exception
{
public:
	ExcelException(const String& message)
		: Exception(message)
	{}

	ExcelException(const char* message)
		: Exception(message)
	{}

protected:
	ExcelException(const String& classWhere, const String& method, const String& message)
		: Exception(classWhere + "::" + method + ": " + message)
	{}

	ExcelException(const char* classWhere, const char* method, const char* message)
		: Exception(String(classWhere) + "::" + String(method) + ": " + String(message))
	{}
};

//---------------------------------------------------------------------------
#endif

