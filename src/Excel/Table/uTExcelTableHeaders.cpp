//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTableHeaders.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------

TExcelTableHeaders::TExcelTableHeaders(TExcelObjectRanged* Parent, const Variant& data)
    : TExcelObjectRanged(Parent, data)
{
}

/* TExcelTableHeaders::TExcelTableHeaders(const TExcelTableHeaders&) {
}

TExcelTableHeaders::TExcelTableHeaders(const TExcelCell&) {
} */

TExcelTableHeaders::~TExcelTableHeaders() {
}


//---------------------------------------------------------------------------
}

