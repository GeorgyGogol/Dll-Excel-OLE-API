//---------------------------------------------------------------------------


#pragma hdrstop

#include "uFunctions.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

namespace CorrectInsert {

void InsertIntoVarCell(const Variant& vData, Variant& vCell, const String& sNullValue)
{
	// В зависимости от передаваемого значения

	// Буферы:
	String buf;
	double dBuf; // long double вызывает непонятки у компилятора
	TDateTime dtBuf;

	/// Для начала сконвертируем в строку
	buf = VarToStrDef(vData, sNullValue.c_str());

	if (buf == sNullValue) {
		//vCell.OlePropertySet("Value", buf);
		return;
	}

	if(TryStrToFloat(buf, dBuf)){
		vCell.OlePropertySet("Value", dBuf);
	}
	else if (TryStrToTime(buf, dtBuf)){
		vCell.OlePropertySet("Value", dtBuf);
	}
	else {
		vCell.OlePropertySet("Value", buf);
	}
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

void InsertIntoVarArray(const String& sData, Variant& vArray, unsigned int row, unsigned int col, const String& sNullValue)
{
	InsertIntoVarArray(Variant(sData), vArray, row, col, sNullValue);
}

}



namespace Converters {

String DefineFormat(exl::ExcelFormats format) {
	using namespace exl;

	switch (format)
    {
    case ExcelFormats::Text:
        return "@";
        
    case ExcelFormats::Number:
        return "0.00";
        
    case ExcelFormats::Integer:
        return "0";
        
    case ExcelFormats::DateShort:
        return "m/d/yyyy";
        
    case ExcelFormats::DateLong:
        return "[$-F800]dddd, mmmm dd, yyyy";

    case ExcelFormats::General:
    default:
        return "General";

    }
}

}

