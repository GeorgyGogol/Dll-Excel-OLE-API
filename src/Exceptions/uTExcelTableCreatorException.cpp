//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTableCreatorException.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
ExcelTableCreatorException::ExcelTableCreatorException(const String& method, const String& message)
    : ExcelException("TTableCreator", method, message)
{}

