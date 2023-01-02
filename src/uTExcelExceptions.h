//---------------------------------------------------------------------------

#ifndef uTExcelExceptionsH
#define uTExcelExceptionsH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uDll.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI ExcelException : public Exception
{
public:
	ExcelException(const String& message)
		: Exception(message)
	{}

protected:
	ExcelException(const String& classWhere, const String& method, const String& message)
		: Exception(classWhere + "::" + method + ": " + message)
	{}
};

//---------------------------------------------------------------------------
class DLL_EI ExcelTableCreatorException : public ExcelException
{
public:
	ExcelTableCreatorException(const String& method, const String& message)
		: ExcelException("TTableCreator", method, message)
	{}
};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

