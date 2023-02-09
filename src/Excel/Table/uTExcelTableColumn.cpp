//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTableColumn.h"
#include "uFunctions.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelTableColumn::TExcelTableColumn(TExcelObject* pParent, const Variant& data)
	: TExcelObjectRangedTemplate<TExcelTableColumn>(pParent, data)
	{}

TExcelTableColumn::TExcelTableColumn(TExcelTableColumn& src)
	: TExcelObjectRangedTemplate<TExcelTableColumn>(src)
{}

TExcelTableColumn::~TExcelTableColumn() 
{}

/// @param start Значение с которого начинаем
/// @param step Шаг инкремента
/* TExcelTableColumn* TExcelTableColumn::SetIdentity(int start, int step){
	vData.OlePropertyGet("Range").OlePropertySet("FormulaR1C1", "=R[-1]C+1");
	vData.OlePropertyGet("Range").OlePropertyGet("Item", 2).OlePropertySet("FormulaR1C1", "=1");
	return this;
} */

IFormatManager<TExcelTableColumn>* TExcelTableColumn::GetFormatInterface()
{
	return static_cast<TExcelTableColumn*>(this);
}

TExcelTableColumn* TExcelTableColumn::SetHorizontalAlign(ExcelTextAlign align) 
{
	checkDataValide();
	vData.OlePropertySet("HorizontalAlignment", (int)align);
    return this;
}

TExcelTableColumn* TExcelTableColumn::SetVerticalAlign(ExcelTextAlign align) 
{
	checkDataValide();
	vData.OlePropertySet("VerticalAlignment", (int)align);
    return this;
}

TExcelTableColumn* TExcelTableColumn::SetTextWrap(bool state)
{
    checkDataValide();
    vData.OlePropertySet("WrapText", state);
    return this;
}

TExcelTableColumn* TExcelTableColumn::SetFormat(ExcelFormats format) 
{
    SetFormat(Converters::DefineFormat(format));
    return this;
}

TExcelTableColumn* TExcelTableColumn::SetFormat(const String& format) 
{
    checkDataValide();
    vData.OlePropertySet("NumberFormat", System::StringToOleStr(format));
    return this;
}



}

