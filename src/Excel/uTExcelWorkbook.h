//---------------------------------------------------------------------------

#ifndef uTExcelWorkbookH
#define uTExcelWorkbookH

//---------------------------------------------------------------------------
#include "uTExcelSheet.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelWorkbook : public TExcelObjectTemplate<TExcelWorkbook>
                            , public ITExcelNames
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

    TExcelSheet* GetSheet(const String& sheetName);
    TExcelSheet* GetSheet(unsigned int N);

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
    TExcelWorkbook* Save(const String filePath = "");

	TExcelNameItem* GetNameItem(const String& itemName);
	TExcelNameItem* GetNameItem(unsigned int N);
	TExcelNameItem* AddNamedItem(const String& itemName);
};

}
//---------------------------------------------------------------------------
#endif

