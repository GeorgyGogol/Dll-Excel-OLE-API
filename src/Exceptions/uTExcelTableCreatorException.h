//---------------------------------------------------------------------------

#ifndef uTExcelTableCreatorExceptionH
#define uTExcelTableCreatorExceptionH

//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
class ExcelTableCreatorException : public ExcelException
{
public:
	ExcelTableCreatorException(const String& method, const String& message)
		: ExcelException("TTableCreator", method, message)
	{}
};

#endif

