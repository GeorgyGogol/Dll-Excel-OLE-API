//---------------------------------------------------------------------------

#ifndef uTExcelObjectDataH
#define uTExcelObjectDataH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelObjectNode.h"
//---------------------------------------------------------------------------
namespace exl {
class DLL_EI TExcelObjectData
{
public:
    TExcelObjectData();
    TExcelObjectData(const Variant& data);
    TExcelObjectData(const TExcelObjectData&);
	~TExcelObjectData();
    //void operator=();

protected:
    Variant vData;
    Variant vDataChild;

    // ����-������� ������� ������� ��� ����������� �������������
	unsigned int getChildCountByType(const String& oType);
	void seekAndSetDataChild(const String& oType, const String& name);
    void seekAndSetDataChild(const String& oType, unsigned int Num);

public:
    // ��� ��� ������ ����������, ������� ��� � ������� ��������
    Variant getVariant();

    void Show();
    void Hide();

};

}
//---------------------------------------------------------------------------
#endif

