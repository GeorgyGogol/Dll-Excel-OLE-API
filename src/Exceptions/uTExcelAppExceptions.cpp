//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelAppExceptions.h"

//---------------------------------------------------------------------------
#define ERROR_EXCEL_NOT_CREATED "Excel app is not Created or Attached!"
#define ERROR_EXCEL_CREATED "Cannot create or attach one more Excel in this object! Please, use one TExcelApp to one ExcelInstance"
//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------

ExcelAppException::ExcelAppException(const char* method, const char* message)
		: ExcelException("TExcelApp", method, message)
{}

ExcelAppNotAttachedException::ExcelAppNotAttachedException(const char* method)
    : ExcelAppException(method, ERROR_EXCEL_NOT_CREATED)
{}

ExcelAppAttachedException::ExcelAppAttachedException(const char* method)
    : ExcelAppException(method, ERROR_EXCEL_CREATED)
{}

