//---------------------------------------------------------------------------

#ifndef uTExcelObjectNodeH
#define uTExcelObjectNodeH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../../uTExcelEnums.h"
#include <list>
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObjectNode 
{
public:
protected:
    TExcelObjectNode();
    TExcelObjectNode(TExcelObjectNode* pParent);
	TExcelObjectNode(const TExcelObjectNode& src);
	~TExcelObjectNode();

private:
	TExcelObjectNode* Parent;
	std::list<TExcelObjectNode*> Childs;

	void AddChildClass(TExcelObjectNode* child);
	void RemoveChildClass(TExcelObjectNode* child);

protected:
	TExcelObjectNode* getParentNode() const;
    
};

}
//---------------------------------------------------------------------------
#endif

