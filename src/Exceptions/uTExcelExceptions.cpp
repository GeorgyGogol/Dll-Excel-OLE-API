//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelExceptions.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
ExcelException::ExcelException(const String& message)
    : Exception(message)
{}

ExcelException::ExcelException(const char* message)
    : Exception(message)
{}

ExcelException::ExcelException(const String& classWhere, const String& method, const String& message)
    : Exception(classWhere + "::" + method + ": " + message)
{}

ExcelException::ExcelException(const char* classWhere, const char* method, const char* message)
    : Exception(String(classWhere) + "::" + String(method) + ": " + String(message))
{}

