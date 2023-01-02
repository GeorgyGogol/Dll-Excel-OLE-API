//---------------------------------------------------------------------------

#ifndef uTExcelObjectRangedH
#define uTExcelObjectRangedH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelObject.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObjectRanged : public TExcelObject
{
public:
	//TExcelObjectRanged();
	TExcelObjectRanged(TExcelObject* pParent, const Variant& data);
	TExcelObjectRanged(const TExcelObjectRanged& src);
	~TExcelObjectRanged();

protected:
    AnsiString ColToStrA(unsigned int ACol);
    AnsiString GetRangeString(
        unsigned int startColumn, unsigned int startRow,
        unsigned int endColumn, unsigned int endRow
            );
    AnsiString GetCellString(unsigned int col, unsigned int row);

    //TExcelObject* selectChild(const String& childName);

public:
	// Выбрать и вернуть
	TExcelObjectRanged* select(unsigned int col, unsigned int row);
	TExcelObjectRanged* selectRange(
        unsigned int startColumn, unsigned int startRow,
        unsigned int endColumn, unsigned int endRow
    );

};


//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

