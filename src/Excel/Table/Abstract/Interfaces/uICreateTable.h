#ifndef UICREATETABLE_H
#define UICREATETABLE_H
/** @file
 * @brief Содержит Интерфейс создания таблиц
 */

//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Интерфейс для создания таблиц
 * @ingroup Interfaces
 * 
 * Отвечает за наличие всевозможных методов создания новых таблиц из объекта
 * с данными.
 *
 * ---------------------------------------------------------------------- **/
template<class Ttable, class TdataFrom>
class ICreateTable
{
public:
    // @brief Сделать вставку в определенное место и создать таблицу
    // @{
    // @param startColumn Начальный столбик
    // @param startRow Начальная строка
    // @param dataSet Источник данных
    // @param tableTitle Заголовок таблицы
    // @param tableName Имя, по которому можно будет найти таблицу
    // @return Таблица из мира Excel
    virtual Ttable* CreateTable(unsigned int startColumn, unsigned int startRow, TdataFrom* dataSet, const String& tableTitle, const String& tableName) = 0;
    virtual Ttable* CreateTable(unsigned int startColumn, unsigned int startRow, TdataFrom* dataSet, const String& tableTitle) = 0;
    virtual Ttable* CreateTable(unsigned int startColumn, unsigned int startRow, TdataFrom* dataSet) = 0;

    // @brief Сделать вставку в начало и создать таблицу
    // @{
    // @param dataSet Источник данных
    // @param tableTitle Заголовок таблицы
    // @param tableName Имя, по которому можно будет найти таблицу
	inline Ttable* CreateTable(TdataFrom* dataSet, const String& tableTitle, const String& tableName)
		{ return CreateTable(1, 1, dataSet, tableTitle, tableName); }
	inline Ttable* CreateTable(TdataFrom* dataSet, const String& tableTitle)
		{ return CreateTable(1, 1, dataSet, tableTitle); }
	inline Ttable* CreateTable(TdataFrom* dataSet)
		{ return CreateTable(1, 1, dataSet); }
    
};

}

//---------------------------------------------------------------------------
#endif // UICREATETABLE_H
