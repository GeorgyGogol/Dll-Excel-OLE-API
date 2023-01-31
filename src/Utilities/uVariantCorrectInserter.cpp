//---------------------------------------------------------------------------


#pragma hdrstop

#include "uVariantCorrectInserter.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

namespace CorrectInsert {

void InsertIntoSingleVariant(const Variant& vData, Variant& vCell, const String& sNullValue)
{
	/*
	String buf;
	long double dBuf;
	TDateTime dtBuf;

	buf = VarToStrDef(vData, sNullValue.c_str());

	if(TryStrToFloat(buf, dBuf)){
		vCell.OlePropertySet("Value", dBuf);
	}
	else if (TryStrToTime(buf, dtBuf)){
		vCell.OlePropertySet("Value", dtBuf);
	}
	else {

	}
	*/

	vCell.OlePropertySet("Value", vData);
}

void InsertIntoVarArray(const Variant& vData, Variant& varArr, unsigned int row, unsigned int col, const String& sNullValue)
{
	String buf;
	long double dBuf;
	TDateTime dtBuf;

	buf = VarToStrDef(vData, sNullValue.c_str());

	if(TryStrToFloat(buf, dBuf)){
		varArr.PutElement(dBuf, row, col);
	}
	else if (TryStrToTime(buf, dtBuf)){
		varArr.PutElement(dtBuf, row, col);
	}
	else {
		varArr.PutElement(buf, row, col);
	}
}

}

