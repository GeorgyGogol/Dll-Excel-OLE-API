//---------------------------------------------------------------------------

#ifndef uTExcelTableCreatorExceptionH
#define uTExcelTableCreatorExceptionH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI ExcelTableCreatorException : public ExcelException
{
public:
	ExcelTableCreatorException(const String& method, const String& message)
		: ExcelException("TTableCreator", method, message)
	{}
};


}
#endif

