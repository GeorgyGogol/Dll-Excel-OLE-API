//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObjectNode.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelObjectNode::TExcelObjectNode()
    : Parent(NULL)
{
}

TExcelObjectNode::TExcelObjectNode(TExcelObjectNode* pParent) 
    : Parent(pParent)
{
    Parent->AddChildClass(this);
}

TExcelObjectNode::TExcelObjectNode(const TExcelObjectNode& src) 
{
    Parent = src.Parent;
    Parent->AddChildClass(this);
}

TExcelObjectNode::~TExcelObjectNode() 
{
	/**
	 * Финальный шаг удаления всех объектов в иерархии - уничтожить
	 * все свои "дочерние" элементы и вычеркнуть себя у родителя. 
	 * Вниз - удаляем всех, вверх - просто сообщаем, что объекта больше
	 * нет. Основная реализация деструктора полностью ложится на дочерний 
	 * класс (в иерархии наследования, НЕ принадлежности)
	 */
	
	while (Childs.size() > 0) // Пока есть 
	{
		/// \details При удалении объект сам себя вычеркнет, т.е.
		/// нам нужон всегда только первый элемент.
		/// Т.о. освобождение памяти происходит рекурсивно.
		delete *Childs.begin(); // Удоляем
	}

	/// @details Если есть родитель - сообщим ему "новость" и забудем про него
	if (Parent) {
		Parent->RemoveChildClass(this);
		Parent = 0;
	}
}

void TExcelObjectNode::AddChildClass(TExcelObjectNode* child) 
{
    Childs.push_back(child);
}

void TExcelObjectNode::RemoveChildClass(TExcelObjectNode* child)
{
    Childs.remove(child);
}

TExcelObjectNode* TExcelObjectNode::getParentNode() const
{
	return Parent;
}

#ifdef ENABLE_USAGE_STATISTIC
/// @return Итератор на первый элемент контейнера
std::list<TExcelObjectNode*>::iterator TExcelObjectNode::Begin()
{
	return Childs.begin();
}

/// @return Итератор на конец контейнера
std::list<TExcelObjectNode*>::iterator TExcelObjectNode::End()
{
    return Childs.end();
}
#endif

}

