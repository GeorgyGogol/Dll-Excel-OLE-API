//---------------------------------------------------------------------------

#ifndef uTExcelObjectDataH
#define uTExcelObjectDataH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelObjectNode.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelObjectData
{
public:
protected:
    TExcelObjectData();
    TExcelObjectData(const Variant& data);
	TExcelObjectData(const TExcelObjectData&);
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

