//---------------------------------------------------------------------------

#ifndef uFunctionsH
#define uFunctionsH

//---------------------------------------------------------------------------

#include "Classes.hpp"
#include "uEnums.h"

namespace CorrectInsert {
    
    void InsertIntoVarCell(const Variant& vData, Variant& vCell, const String& sNullValue = "");
    void InsertIntoVarArray(const Variant& vData, Variant& vArray, unsigned int row, unsigned int col, const String& sNullValue = "");
    void InsertIntoVarArray(const String& vData, Variant& vArray, unsigned int row, unsigned int col, const String& sNullValue = "");

}

namespace Converters {

    String DefineFormat(exl::ExcelFormats format);

}

#endif
