//---------------------------------------------------------------------------

#ifndef uTExcelObjectH
#define uTExcelObjectH

//---------------------------------------------------------------------------
#include "SubClasses/uTExcelObjectData.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObject : public TExcelObjectNode<TExcelObject>, public TExcelObjectData
{
public:
	TExcelObject();
	TExcelObject(TExcelObject* Parent, const Variant& data);
	TExcelObject(const TExcelObject& src);
	~TExcelObject();

protected:

public:
	Variant getParentVariant();
	TExcelObject* getCurrentSelectedChild();

};
//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif
