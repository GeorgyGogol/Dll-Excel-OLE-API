#ifndef uTTableCreatorH
#define uTTableCreatorH
/** @file Создатель таблиц
 * @brief Отвечает за создание таблиц
 */

#include "uTGridHeaders.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс Рабочей книги
 * @ingroup ExcelUtilities 
 * 
 * Класс-помощник (фабрика) создания таблиц на основе DataSet-ов
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TTableCreator
{
public:
    //TTableCreator();
	TTableCreator(TExcelObject* pSheet, TDataSet* dataSet, const String& tableTitle, const String& tableName);
	TTableCreator(TExcelObject* pSheet, TDataSet* dataSet, const String& tableTitle);
	TTableCreator(TExcelObject* pSheet, TDataSet* dataSet);

	TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh, const String& tableTitle, const String& tableName);
	TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh, const String& tableTitle);
	TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh);

    ~TTableCreator();
    
private:
	TExcelObject* Sheet; ///< Страница, на которой будет все происходить
    TGridHeaders* Headers; ///< Заголовки
    Variant varData; ///< Содержимое, уже подготовленное
    unsigned int nRecords; ///< Количество записей

    void readDataSet(TDataSet* dataSet);

    void init(TDataSet* dataSet); ///< Иницилизация с помощью дата-сета
    void init(TDBGridEh* gridEh); ///< Иницилизация с помощью GridEh-а

    void check(unsigned int& col, unsigned int& row); ///< Проверка столбика и строчки

public:
    String Title; ///< Заголовок
    String TableName; ///< Название таблицы
    String TableStyle; ///< Стиль таблицы

    void ResetData(); ///< очистить данные

    void PrepareNewData(TExcelObject* sheet, TDataSet* dataSet, const String& tableTitle = "", const String& tableName = "");
    void PrepareNewData(TExcelObject* sheet, TDBGridEh* gridEh, const String& tableTitle = "", const String& tableName = "");

    bool CanCreate() const; ///< Можем ли мы создать таблицу

    TExcelCells* InsertData(unsigned int col, unsigned int row, bool needInsertFieldNames = false); ///< Вставить прочитанные данные на лист
    TExcelTable* CreateTable(unsigned int col, unsigned int row); ///< Создать таблицу от заданной точки

};

}

//---------------------------------------------------------------------------
#endif


