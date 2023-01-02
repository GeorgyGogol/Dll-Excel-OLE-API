//---------------------------------------------------------------------------

#ifndef uTExcelCellH
#define uTExcelCellH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "Abstract/uTExcelObjectRanged.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelCell : public TExcelObjectRanged {
public:
    //TExcelCell();
    TExcelCell(TExcelObjectRanged* pParent, const Variant& data);
    TExcelCell(const TExcelCell&);
    ~TExcelCell();

private:
	//TExcelObject* getCurrentSelectedChild();

public:
    TExcelCell* Insert(const Variant& data);
    TExcelCell* InsertString(const String& data);
    TExcelCell* InsertFormula(String& formula);
    
    TExcelCell* Merge(); // Объеденить ячейки
    TExcelCell* SetHorizontalAlign(ExcelTextAlign align);
    TExcelCell* SetVerticalAlign(ExcelTextAlign align);
    
    String ReadString();
};

//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif

