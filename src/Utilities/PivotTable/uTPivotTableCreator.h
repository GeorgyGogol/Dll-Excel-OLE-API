//---------------------------------------------------------------------------

#ifndef uTPivotTableCreatorH
#define uTPivotTableCreatorH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTPivotSettings.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TPivotTableCreator
{
public:
    TPivotTableCreator(TExcelTable* source);
    //TPivotTableCreator(TExcelTable book, const String& tableName);
    ~TPivotTableCreator();
    
private:
    TExcelTable* SourceTable;

    TExcelTable* initPivot(
        TExcelObject* aimSheet,
        unsigned int col, unsigned int row, 
        TPivotSettings* PivotSettings
    );

public:
    //void ChangeSheet(TExcelObject* sheet);
    void ChangeSourceTable(TExcelTable* sourceNew);
    //void ChangeSourceTable(const String& sourceTableName);
    //TExcelTable* CreateTable(TExcelObject* sheet, unsigned int col, unsigned int row, TPivotSettings& PivotSettings);
    TExcelTable* CreateTable(TExcelObject* aimSheet, unsigned int col, unsigned int row, TPivotSettings* PivotSettings);

};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

