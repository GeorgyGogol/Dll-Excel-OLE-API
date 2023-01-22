//---------------------------------------------------------------------------

#ifndef uTExcelObjectDataH
#define uTExcelObjectDataH

//---------------------------------------------------------------------------
#include "uDll.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObjectData
{
public:
    TExcelObjectData();
    TExcelObjectData(const Variant& data);
	TExcelObjectData(const TExcelObjectData&);
protected:
	~TExcelObjectData();
    //void operator=();

protected:
    Variant vData;
    Variant vDataChild;

    void checkDataValide();

    // Обще-типовые обертки функций для внутреннего использования
	unsigned int getChildCountByType(const String& oType);
	void seekAndSetDataChild(const String& oType, const String& name);
    void seekAndSetDataChild(const String& oType, unsigned int Num);

public:
    // Это для особых извращений, которых нет в типовых решениях
	Variant getVariant();

};

}
//---------------------------------------------------------------------------
#endif

