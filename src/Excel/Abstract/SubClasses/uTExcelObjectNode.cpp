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

/// @param src Источник копии
TExcelObjectNode::TExcelObjectNode(const TExcelObjectNode& src) 
{
	/**
	 * При создании копии элемена, мы должны запомнить общего родителя
	 * и сообщить ему радостную весть.
	 */
    Parent = src.Parent;
    Parent->AddChildClass(this);
}

/// @param src Источник копии
/// @return указатель на левый объект
TExcelObjectNode TExcelObjectNode::operator=(const TExcelObjectNode& src)
{
	if (Parent) Parent->RemoveChildClass(this);
    Parent = src.Parent;
    Parent->AddChildClass(this);
	return this;
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
	ClearChilds();

	/// Если есть родитель - сообщим ему "новость" и забудем про него
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

void TExcelObjectNode::ClearChilds()
{
	/**
	 * Если ВДРУГ у нас появилась необходимость удалить всю дочернюю
	 * структуру, то этот метод сделает это! Метод переехал из действия
	 * деструктора.
	 */
	while (Childs.size() > 0) // Пока есть дочерние элементы
	{
		/** 
		 * При удалении объект сам себя вычеркнет, т.е. нам нужон всегда только
		 * первый элемент. Т.о. освобождение памяти происходит лавиной.
		 */
		delete *Childs.begin(); // Удоляем
	}
}

/// @{ @addtogroup ENABLE_USAGE_STATISTIC
/// @{
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
/// @} @}

}

