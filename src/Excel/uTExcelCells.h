//---------------------------------------------------------------------------

#ifndef uTExcelCellsH
#define uTExcelCellsH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelCell.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelCells : public TExcelCell
{
public:
    TExcelCells(TExcelObjectRanged* pParent, const Variant& data);
    TExcelCells(TExcelCells&);
    ~TExcelCells();

public:
    unsigned int GetColumnsCount();
    unsigned int GetRowCount();

    TExcelCell* GetCell(unsigned int col, unsigned int row);
    TExcelCells* GetCells(
        unsigned int startColumn, unsigned int startRow,
        unsigned int endColumn, unsigned int endRow
    );
};


//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif

