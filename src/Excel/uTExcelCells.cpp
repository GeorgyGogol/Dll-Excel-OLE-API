//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelCells.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelCells::TExcelCells(TExcelObjectRanged* pParent, const Variant& data) 
    : TExcelCell(pParent, data)
{}

TExcelCells::TExcelCells(TExcelCells& src) 
    : TExcelCell(src)
{}

TExcelCells::~TExcelCells() 
{}

unsigned int TExcelCells::GetColumnsCount() 
{
    return 0;
}

unsigned int TExcelCells::GetRowCount() 
{
    return 0;
}

TExcelCell* TExcelCells::GetCell(unsigned int col, unsigned int row) 
{
    return 0;
}

TExcelCells* TExcelCells::GetCells(
    unsigned int startColumn, unsigned int startRow,
    unsigned int endColumn, unsigned int endRow
) 
{
    return 0;
}

}

