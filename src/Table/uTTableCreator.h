//---------------------------------------------------------------------------

#ifndef uTTableCreatorH
#define uTTableCreatorH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTGridHeaders.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TTableCreator
{
public:
    TTableCreator();
    TTableCreator(TExcelObjectRanged* sheet, TDataSet* dataSet, const String& tableTitle, const String& tableName);
    TTableCreator(TExcelObjectRanged* sheet, TDataSet* dataSet, const String& tableTitle);
    TTableCreator(TExcelObjectRanged* sheet, TDataSet* dataSet);

	TTableCreator(TExcelObjectRanged* sheet, TDBGridEh* gridEh, const String& tableName, const String& tableTitle);
    ~TTableCreator();
    
private:
	TExcelObjectRanged* Sheet;
    TGridHeaders* Headers;
    Variant varData;
    unsigned int nRecords;

    void readDataSet(TDataSet* dataSet);

    void init(TDataSet* dataSet);
    void init(TDBGridEh* gridEh);

    void check(unsigned int& col, unsigned int& row);

public:
    String TableName;
    String TableStyle;
    String Title;

    void ResetData();

    void PrepareNewData(TExcelObjectRanged* sheet, TDataSet* dataSet, const String& tableTitle = "", const String& tableName = "");
    void PrepareNewData(TExcelObjectRanged* sheet, TDBGridEh* gridEh, const String& tableTitle = "", const String& tableName = "");

    bool CanCreate() const;

    TExcelCells* InsertData(unsigned int col, unsigned int row, bool needInsertFieldNames = false);
    //TExcelTable* InsertTable(unsigned int col, unsigned int row);
    TExcelTable* CreateTable(unsigned int col, unsigned int row);

};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif


