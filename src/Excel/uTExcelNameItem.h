//---------------------------------------------------------------------------

#ifndef uTExcelNameItemH
#define uTExcelNameItemH

//---------------------------------------------------------------------------
#include "uTExcelObjectTemplate.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelNameItem : public TExcelObjectTemplate<TExcelNameItem>
{
private:
    //TExcelNameItem();

public:
    TExcelNameItem(TExcelObject* pParent, const Variant& data);
    TExcelNameItem(const TExcelNameItem& src);
    virtual ~TExcelNameItem();

public:
	TExcelNameItem* SetValue(const Variant& val);
	TExcelNameItem* SetValue(const String& val);
	TExcelNameItem* SetRefers(const String& address);

};

class ITExcelNames {
public:
	virtual TExcelNameItem* GetNameItem(const String& itemName) = 0;
	virtual TExcelNameItem* GetNameItem(unsigned int N) = 0;

	virtual TExcelNameItem* AddNamedItem(const String& itemName) = 0;

};

}
//---------------------------------------------------------------------------
#endif
