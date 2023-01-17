//---------------------------------------------------------------------------

#ifndef uTExcelAppExceptionsH
#define uTExcelAppExceptionsH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
#define ERROR_EXCEL_NOT_CREATED "Excel app is not Created or Attached!"
#define ERROR_EXCEL_CREATED "Cannot create or attach one more Excel in this object! Please, use one TExcelApp to one ExcelInstance"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI ExcelAppException : public ExcelException
{
public:
	ExcelAppException(const char* method, const char* message)
		: ExcelException("TExcelApp", method, message)
	{}
};
//---------------------------------------------------------------------------
class DLL_EI ExcelAppNotAttachedException : public ExcelAppException
{
public:
	ExcelAppNotAttachedException(const char* method)
		: ExcelAppException(method, ERROR_EXCEL_NOT_CREATED)
	{}
};

//---------------------------------------------------------------------------
class DLL_EI ExcelAppAttachedException : public ExcelAppException
{
public:
	ExcelAppAttachedException(const char* method)
		: ExcelAppException(method, ERROR_EXCEL_CREATED)
	{}
};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#undef ERROR_EXCEL_NOT_CREATED
#undef ERROR_EXCEL_CREATED
//---------------------------------------------------------------------------
#endif

