#ifndef uTExcelObjectNodeH
#define uTExcelObjectNodeH
/** @file
 * @ingroup Abstract
 * @brief Позиция объекта в этой кастовой иерархии
 */

#include "uDll.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Базовый класс для выстраивания иерархии получаемых объектов
 * @ingroup Abstract
 *
 * Является абстрактным понятием.
 * 
 * Зона ответственности: связывание объектов, выстраивание иерархии что-где 
 * находится и кому принадлежит. При создании элемента, конструктор этого
 * класса сохраняет ссылки на родителя (кроме хост-приложения, у него нет 
 * родительского элемента), а также все сгенерированные дочерние элементы.
 * 
 * Не имеет открытых методов, так как является чисто служебным элементом.
 * Не доступен для создания.
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelObjectNode
{
//public:
protected:
    TExcelObjectNode(); ///< Базовый, без родителя
    TExcelObjectNode(TExcelObjectNode* pParent); ///< Для дочернего объекта
	TExcelObjectNode(const TExcelObjectNode& src); ///< Копирование
	~TExcelObjectNode(); ///< Удалить объект

	/// Операция присваивания
	TExcelObjectNode operator=(const TExcelObjectNode& src);

private:
	TExcelObjectNode* Parent; ///< Указатель на родителя
	std::list<TExcelObjectNode*> Childs; ///< Список дочерних элементов

	void AddChildClass(TExcelObjectNode* child); ///< Добавить дочерний класс
	void RemoveChildClass(TExcelObjectNode* child); ///< Убрать дочерний класс

protected:
	TExcelObjectNode* getParentNode() const; ///< Получить указатель на родителя

	void ClearChilds(); ///< Удалить все дочерние элементы

#ifdef ENABLE_USAGE_STATISTIC
	std::list<TExcelObjectNode*>::iterator Begin(); ///< Итератор на первого ребенка
	std::list<TExcelObjectNode*>::iterator End(); ///< Итератор на конец массива с детьми
#endif
};

}

//---------------------------------------------------------------------------
#endif

