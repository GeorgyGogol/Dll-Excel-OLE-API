#ifndef uACreateTableControllerH
#define uACreateTableControllerH
/** @file
 * @brief Содержит абстрактный контроллер для порождения таблиц
 */

#include "uTExcelTable.h"
#include "uICreateTable.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Абстрактный класс для добавления полного функционала создания таблиц
 * @ingroup Abstract
 * 
 * Абстракция, которая собирает все возможные методы получения объектов
 * таблиц.
 * 
 * Включает тот необходимый стандартный функционал "а как и из чего можно
 * собрать таблицу". Словом, сюда все что нужно для создания новой таблицы
 * было слито сюда.
 * 
 * ---------------------------------------------------------------------- **/
class ACreateTableController :
	public ICreateTable<TExcelTable, TDataSet>,
	public ICreateTable<TExcelTable, TDBGridEh>
{
protected:
    ACreateTableController(); ///< Сюда иницилизация элементов управления

protected:
    bool NeedDisableDataSet; ///< Нужно ли выключать датасет

public:
    void setNeedDisableDataSet(bool isNeedDisable); ///< Управление необходимостью вкл/выкл ДатаСета

	// @brief Создать таблицу в/на диапазоне
	// @param startColumn Колонка от которой (включительно)
	// @param startRow Стро
	// @return 
	//virtual TExcelTable* CreateTable(unsigned int startColumn, unsigned int startRow) = 0;
	//virtual TExcelTable* Create

};

}

//---------------------------------------------------------------------------
#endif
