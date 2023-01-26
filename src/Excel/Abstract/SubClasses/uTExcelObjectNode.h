#ifndef uTExcelObjectNodeH
#define uTExcelObjectNodeH

#include "uDll.h"
//---------------------------------------------------------------------------
namespace exl {
/** @addtogroup Abstract
 * @{
 * @brief Базовый класс для выстраивания иерархии получаемых объектов
 * 
 * Является абстрактным понятием.
 * 
 * Зона ответственности: связывание объектов, выстраивание иерархии что-где 
 * находится и кому принадлежит. При создании элемента, конструктор этого
 * класса сохраняет ссылки на родителя (кроме хост-приложения, у него нет 
 * родительского элемента), а также все сгенерированные дочерние элементы.
 * 
 * Не имеет открытых методов, так как является чисто служебным элементом.
 */
class DLL_EI TExcelObjectNode
{
public:
    TExcelObjectNode();
    TExcelObjectNode(TExcelObjectNode* pParent);
	TExcelObjectNode(const TExcelObjectNode& src);
protected:
	~TExcelObjectNode();

private:
	TExcelObjectNode* Parent; ///< Указатель на родителя
	std::list<TExcelObjectNode*> Childs; ///< Список дочерних элементов

	void AddChildClass(TExcelObjectNode* child);
	void RemoveChildClass(TExcelObjectNode* child);

protected:
	TExcelObjectNode* getParentNode() const;

#ifdef ENABLE_USAGE_STATISTIC
	/// @brief Итератор на первого ребенка
	std::list<TExcelObjectNode*>::iterator Begin();
	/// @brief Итератор на конец массива с детьми
	std::list<TExcelObjectNode*>::iterator End(); 
#endif
};

}
/// @}
#endif

