//---------------------------------------------------------------------------

#ifndef uTPivotTableCreatorH
#define uTPivotTableCreatorH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
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
        TExcelObjectRanged* Sheet, 
        unsigned int col, unsigned int row, 
        TPivotSettings* PivotSettings
    );

public:
    //void ChangeSheet(TExcelObjectRanged* sheet);
    void ChangeSourceTable(TExcelTable* sourceNew);
    //void ChangeSourceTable(const String& sourceTableName);
    //TExcelTable* CreateTable(TExcelObjectRanged* sheet, unsigned int col, unsigned int row, TPivotSettings& PivotSettings);
    TExcelTable* CreateTable(unsigned int col, unsigned int row, TPivotSettings* PivotSettings);

};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

