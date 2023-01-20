//---------------------------------------------------------------------------

#ifndef uTExcelObjectH
#define uTExcelObjectH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "SubClasses/uTExcelObjectData.h"
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

	// TExcelObject* AddItem(const String& itemName, const Variant& data);
	void* AddNamedItem(const String& itemName);
	void* AddNamedItem(const String& itemName, const String& formula);
	void* AddNamedItem(const String& itemName, const String& formula, const String& comment);
	
	void* SetItem(const String& itemName, const String& data);
	void* SetItem(const String& itemName, const Variant& data);
	void* SetItem(unsigned int i, const String& data);
	void* SetItem(unsigned int i, const Variant& data);
	//TExcelObject* SetItem(const String& itemName, const String& data);

	//TExcelObject* GetItem(unsigned int i);
	//TExcelObject* GetItemAsString(const String& itemName);

	String GetName();

	//String GetNamedObject();

};

/* class DLL_EI TExcelName : public TExcelObjectNode, public TExcelObjectData
{
public:
	TExcelName(TExcelObject* pParent, const Variant& data);
	TExcelName(const TExcelObject& src);

} */

}
//---------------------------------------------------------------------------
#endif


