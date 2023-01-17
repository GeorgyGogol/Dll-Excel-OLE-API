//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTableHeaders.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------

TExcelTableHeaders::TExcelTableHeaders(TExcelObject* Parent, const Variant& data)
    : TExcelCells(Parent, data)
{
}

TExcelTableHeaders::TExcelTableHeaders(TExcelObject* Source)
    : TExcelCells(Source->GetParent(), Source->getVariant())
{}

TExcelTableHeaders::~TExcelTableHeaders() 
{
}

unsigned int TExcelTableHeaders::HeadersDepth(){
    return GetRowCount();
}

unsigned int TExcelTableHeaders::HeadersCount() {
    return GetColumnsCount();
}


String TExcelTableHeaders::GetTitle(unsigned int i) {
    String out;
    TExcelCells* select = GetCell(i + 1, 1);
    out = select->ReadValueString();
    delete select;
    return out;
}

void TExcelTableHeaders::SetTitle(unsigned int i, const String& title) {
    TExcelCells* select = GetCell(i + 1, 1);
    select->InsertString(title);
    delete select;
}


//---------------------------------------------------------------------------
}

