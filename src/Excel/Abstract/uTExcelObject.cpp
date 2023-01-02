//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObject.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelObject::TExcelObject()
	: TExcelObjectNode<TExcelObject>(), TExcelObjectData()
{
}

TExcelObject::TExcelObject(TExcelObject* pParent, const Variant& data)
    : TExcelObjectNode<TExcelObject>(pParent), TExcelObjectData(data)
{
}

TExcelObject::TExcelObject(const TExcelObject& src)
	: TExcelObjectNode<TExcelObject>(src), TExcelObjectData(src)
{
}

TExcelObject::~TExcelObject()
{
}

Variant TExcelObject::getParentVariant(){
	return ((TExcelObjectData*)Parent)->getVariant();
}

TExcelObject* TExcelObject::getCurrentSelectedChild()
{
	TExcelObject* out = new TExcelObject(this, vDataChild);
	return out;
}

}


