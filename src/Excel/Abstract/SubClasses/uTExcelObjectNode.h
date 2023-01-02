//---------------------------------------------------------------------------

#ifndef uTExcelObjectNodeH
#define uTExcelObjectNodeH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "Heads/uDll.h"
#include <list>
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
template<class TNode> class DLL_EI TExcelObjectNode {
public:
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
		// ��������� ��� �������� ���� �������� � �������� -
		// ���������� ��� ���� "��������" �������� � ����������
		// ���� � ��������. ���� - ������� ����, ����� - ������
		// ��������, ��� ������� ������ ���.
		// �������� ���������� ����������� ��������� �������
		// �� �������� �����.

		while (Childs.size() > 0) // ���� ����
		{
			// ��� �������� ������ ��� ���� ���������, �.�.
			// ��� ����� ������ ������ ������ �������.
			// �.�. ������������ ������ ���������� ����������.
			delete *Childs.begin(); // �������
		}

		if (Parent) { // ���� ���� �������� -
			Parent->RemoveChildClass(this); // ������� ��� �������
			Parent = 0; // ������� ��� ����
		}
	}

protected:
	TExcelObjectNode* Parent;
    std::list<TExcelObjectNode*> Childs;


	void TExcelObjectNode::AddChildClass(TExcelObjectNode* child)
	{
		Childs.push_back(child);
	}

	void TExcelObjectNode::RemoveChildClass(TExcelObjectNode* child)
	{
		Childs.remove(child);
	}

public:
	TNode* TExcelObjectNode::getParent() const {
		return (TNode*)Parent;
	}

};
//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

