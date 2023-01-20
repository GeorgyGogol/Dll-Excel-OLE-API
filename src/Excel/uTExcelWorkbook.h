//---------------------------------------------------------------------------

#ifndef uTExcelWorkbookH
#define uTExcelWorkbookH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelSheet.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelWorkbook : public TExcelObjectTemplate<TExcelWorkbook>
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

    TExcelTable* CreateTable(
        TDataSet* dataSet, const String& sheetName,
        const String& tableTitle, const String& tableName,
        bool needDisableSet = false
    );
    TExcelTable* CreateTable(
        TDataSet* dataSet, const String& sheetName, 
        const String& tableTitle,
        bool needDisableSet = false
    );
    TExcelTable* CreateTable(
        TDataSet* dataSet, const String& sheetName,
        bool needDisableSet = false
    );
    TExcelTable* CreateTable(
        TDataSet* dataSet,
        bool needDisableSet = false
    );

    TExcelTable* CreateTable(
        TDBGridEh* gridEh, const String& sheetName,
        const String& tableTitle, const String& tableName,
        bool needDisableSet = false
    );
    TExcelTable* CreateTable(
        TDBGridEh* gridEh, const String& sheetName,
        const String& tableTitle,
        bool needDisableSet = false
    );
    TExcelTable* CreateTable(
        TDBGridEh* gridEh, const String& sheetName,
        bool needDisableSet = false
    );
    TExcelTable* CreateTable(
        TDBGridEh* gridEh,
        bool needDisableSet = false
    );

    TExcelWorkbook* Save();
};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

