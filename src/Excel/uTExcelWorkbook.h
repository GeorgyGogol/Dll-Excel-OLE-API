#ifndef uTExcelWorkbookH
#define uTExcelWorkbookH

#include "uTExcelSheet.h"
#include "uACreateTableController.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс Рабочей книги
 * @ingroup ExcelClientObjects
 * 
 * Проекция книги из мира Excel
 * @todo Норм документация
 * 
 * @note ACreateTableController справедлив для текущего листа
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelWorkbook : 
    public TExcelObjectTemplate<TExcelWorkbook>,
    //public ACreateTableController,
    public ITExcelNames
{
public:
    TExcelWorkbook(TExcelObject* pParent, const Variant& data);
	TExcelWorkbook(const TExcelWorkbook&);
    ~TExcelWorkbook();

public:
    unsigned int SheetCount();
	TExcelSheet* CreateSheet();
	TExcelSheet* CreateSheet(const String& sheetName);
	TExcelSheet* GetCurrentSheet();
    TExcelSheet* SelectSheet(const String& sheetName);
    TExcelSheet* SelectSheet(unsigned int N);
    TExcelSheet* GetSheet(const String& sheetName);
    TExcelSheet* GetSheet(unsigned int N);

    TExcelWorkbook* Save();
    TExcelWorkbook* Save(const String filePath = "");

	// ACreateTableController override
	/* TExcelTable* GetTable(const String& tableName);
	// Вставка с созданием НАСТОЯЩЕЙ ТАБЛИЦЫ как объекта в координаты
 	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet, const String& tableTitle, const String& tableName);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet, const String& tableTitle);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDataSet* dataSet);
	// Без координат
 	TExcelTable* CreateTable(TDataSet* dataSet, const String& tableTitle, const String& tableName);
	TExcelTable* CreateTable(TDataSet* dataSet, const String& tableTitle);
	TExcelTable* CreateTable(TDataSet* dataSet);
	// Тут на основе GridEh
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDBGridEh* gridEh, const String& tableTitle, const String& tableName);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDBGridEh* gridEh, const String& tableTitle);
	TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow, TDBGridEh* gridEh);
	// Без координат
	TExcelTable* CreateTable(TDBGridEh* gridEh, const String& tableTitle, const String& tableName);
	TExcelTable* CreateTable(TDBGridEh* gridEh, const String& tableTitle);
	TExcelTable* CreateTable(TDBGridEh* gridEh); */

	// ITExcelNames override
	TExcelNameItem* GetNameItem(const String& itemName);
	TExcelNameItem* GetNameItem(unsigned int N);
	TExcelNameItem* AddNamedItem(const String& itemName);
};

}

//---------------------------------------------------------------------------
#endif

