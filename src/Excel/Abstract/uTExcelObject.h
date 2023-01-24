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
	//unsigned int GetSize() const;

public:
	TExcelObject* GetParent() const;
	Variant GetParentVariant();

	TExcelObject* GetCurrentSelectedChild();

	TExcelObject* Show();
	TExcelObject* Hide();
	TExcelObject* SetName(const String& newName);

    //TExcelObject* GetNameItem(const String& itemName);
	//TExcelObject* GetNameItem(unsigned int N);

	//TExcelObject* AddNamedItem(const String& itemName);


	// TExcelObject* AddItem(const String& itemName, const Variant& data);
	// void* AddNamedItem(const String& itemName);
	// void* AddNamedItem(const String& itemName, const String& formula);
	// void* AddNamedItem(const String& itemName, const String& formula, const String& comment);
	
	// void* SetItem(const String& itemName, const String& data);
	// void* SetItem(const String& itemName, const Variant& data);
	// void* SetItem(unsigned int i, const String& data);
	// void* SetItem(unsigned int i, const Variant& data);
	//TExcelObject* SetItem(const String& itemName, const String& data);

	//TExcelObject* GetItem(unsigned int i);
	//TExcelObject* GetItemAsString(const String& itemName);

	String GetName();

	//String GetNamedObject();

	unsigned int SizeOfThis();

};

void InsertIntoSingleVariant(const Variant& vData, Variant& vCell, const String& sNullValue = "");
void InsertIntoVarArray(const Variant& vData, Variant& vCell, unsigned int row, unsigned int col, const String& sNullValue = "");

}
//---------------------------------------------------------------------------
#endif

