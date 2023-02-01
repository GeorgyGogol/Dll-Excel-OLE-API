//---------------------------------------------------------------------------

#ifndef uVariantCorrectInserterH
#define uVariantCorrectInserterH

//---------------------------------------------------------------------------

#include "Classes.hpp"

namespace CorrectInsert {
    
    void InsertIntoSingleVariant(const Variant& vData, Variant& vCell, const String& sNullValue = "");
    void InsertIntoVarArray(const Variant& vData, Variant& vCell, unsigned int row, unsigned int col, const String& sNullValue = "");

}

#endif
