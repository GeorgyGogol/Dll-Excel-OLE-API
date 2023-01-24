//---------------------------------------------------------------------------

#ifndef uTExcelObjectNodeH
#define uTExcelObjectNodeH

//---------------------------------------------------------------------------
#include "uDll.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObjectNode
{
public:
    TExcelObjectNode();
    TExcelObjectNode(TExcelObjectNode* pParent);
	TExcelObjectNode(const TExcelObjectNode& src);
protected:
	~TExcelObjectNode();

private:
	TExcelObjectNode* Parent;
	std::list<TExcelObjectNode*> Childs;

	void AddChildClass(TExcelObjectNode* child);
	void RemoveChildClass(TExcelObjectNode* child);

protected:
	TExcelObjectNode* getParentNode() const;

	std::list<TExcelObjectNode*>::iterator Begin();
	std::list<TExcelObjectNode*>::iterator End();
};

}
//---------------------------------------------------------------------------
#endif

