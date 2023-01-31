//---------------------------------------------------------------------------

#ifndef uTExcelObjectH
#define uTExcelObjectH

//---------------------------------------------------------------------------
#include "uDll.h"
#include "uTExcelObjectNode.h"
#include "uTExcelObjectData.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObject : public TExcelObjectNode, public TExcelObjectData
{
public:
	TExcelObject();
	TExcelObject(TExcelObject* pParent, const Variant& data);
	TExcelObject(const TExcelObject& src);
	virtual ~TExcelObject(); // ???

private:

protected:

public:
	TExcelObject* GetParent() const;
	Variant GetParentVariant();

	TExcelObject* GetCurrentSelectedChild();

	TExcelObject* Show();
	TExcelObject* Hide();
	TExcelObject* SetName(const String& newName);

	String GetName();

#ifdef _DEBUG
	virtual unsigned int SizeOfThis();
#endif

};

}
//---------------------------------------------------------------------------
#endif

