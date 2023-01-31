//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObject.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelObject::TExcelObject()
	: TExcelObjectNode(), TExcelObjectData()
{
}

TExcelObject::TExcelObject(TExcelObject* pParent, const Variant& data)
    : TExcelObjectNode(pParent), TExcelObjectData(data)
{
}

TExcelObject::TExcelObject(const TExcelObject& src)
	: TExcelObjectNode(src), TExcelObjectData(src)
{
}

TExcelObject::~TExcelObject()
{
}

TExcelObject* TExcelObject::GetParent() const
{
    return (TExcelObject*)getParentNode();
}

Variant TExcelObject::GetParentVariant() 
{
    TExcelObject* dParent = GetParent();
    if (dParent)
        return dParent->getVariant();
    else return Null();
}

TExcelObject* TExcelObject::GetCurrentSelectedChild()
{
    TExcelObject* out = new TExcelObject(this, vDataChild);
    return out;
}

TExcelObject* TExcelObject::Show() {
    checkDataValide();
    vData.OlePropertySet("Visible", true);
    return this;
}

TExcelObject* TExcelObject::Hide() {
    checkDataValide();
    vData.OlePropertySet("Visible", false);
    return this;
}

TExcelObject* TExcelObject::SetName(const String& newName) {
    checkDataValide();
    vData.OlePropertySet("Name", System::StringToOleStr(newName));
    return this;
}

String TExcelObject::GetName()
{
    checkDataValide();
    Variant name = vData.OlePropertyGet("Name");
    return VarToStr(name);
}

#ifdef _DEBUG
unsigned int TExcelObject::SizeOfThis() {
	unsigned int s = sizeof(*this);

	std::list<TExcelObjectNode*>::iterator it = Begin();

	for (it; it != End(); ++it) {
		s += static_cast<TExcelObject*>(*it)->SizeOfThis();
	}

	return s;
}
#endif


}


