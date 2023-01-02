//---------------------------------------------------------------------------

#ifndef uTExcelWorkbookH
#define uTExcelWorkbookH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelSheet.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelWorkbook : public TExcelObject 
{
public:
	TExcelWorkbook(const TExcelWorkbook&);
    TExcelWorkbook(TExcelObject* pParent, const Variant& data);
    ~TExcelWorkbook();

public:
    //TExcelObject* getExcel();

    unsigned int SheetCount();
	TExcelSheet* CreateSheet();
	TExcelSheet* CreateSheet(const String& sheetName);
	TExcelSheet* GetCurrentSheet();
    TExcelSheet* SelectSheet(const String& sheetName);
    TExcelSheet* SelectSheet(unsigned int N);
};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

