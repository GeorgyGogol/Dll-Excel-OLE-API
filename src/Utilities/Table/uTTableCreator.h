//---------------------------------------------------------------------------

#ifndef uTTableCreatorH
#define uTTableCreatorH

//---------------------------------------------------------------------------
#include "uTGridHeaders.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TTableCreator
{
public:
    TTableCreator();
	TTableCreator(TExcelObject* pSheet, TDataSet* dataSet, const String& tableTitle, const String& tableName);
	TTableCreator(TExcelObject* pSheet, TDataSet* dataSet, const String& tableTitle);
	TTableCreator(TExcelObject* pSheet, TDataSet* dataSet);

	TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh, const String& tableTitle, const String& tableName);
	TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh, const String& tableTitle);
	TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh);

    ~TTableCreator();
    
private:
	TExcelObject* Sheet;
    TGridHeaders* Headers;
    Variant varData;
    unsigned int nRecords;

    void readDataSet(TDataSet* dataSet);

    void init(TDataSet* dataSet);
    void init(TDBGridEh* gridEh);

    void check(unsigned int& col, unsigned int& row);

public:
    String Title;
    String TableName;
    String TableStyle;

    void ResetData();

    void PrepareNewData(TExcelObject* sheet, TDataSet* dataSet, const String& tableTitle = "", const String& tableName = "");
    void PrepareNewData(TExcelObject* sheet, TDBGridEh* gridEh, const String& tableTitle = "", const String& tableName = "");

    bool CanCreate() const;

    TExcelCells* InsertData(unsigned int col, unsigned int row, bool needInsertFieldNames = false);
    TExcelTable* CreateTable(unsigned int col, unsigned int row);

};

}
//---------------------------------------------------------------------------
#endif


