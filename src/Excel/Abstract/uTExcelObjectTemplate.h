//---------------------------------------------------------------------------

#ifndef uTExcelObjectTemplateH
#define uTExcelObjectTemplateH

//---------------------------------------------------------------------------
#include "uTExcelObject.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
template<class T>
class DLL_EI TExcelObjectTemplate : public TExcelObject
{
public:
	TExcelObjectTemplate();
	TExcelObjectTemplate(TExcelObject* pParent, const Variant& data);
	TExcelObjectTemplate(const TExcelObject& src);
protected:
	~TExcelObjectTemplate();

protected:

public:
	T* Show();
	T* Hide();
	T* SetName(const String& newName);

};

}

//---------------------------------------------------------------------------
#endif
