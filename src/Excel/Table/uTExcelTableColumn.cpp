//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelTableColumn.h"

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

TExcelTableColumn::~TExcelTableColumn() {}

TExcelTableColumn* TExcelTableColumn::SetIdentity(int start, int step){
	vData.OlePropertyGet("Range").OlePropertySet("FormulaR1C1", "=R[-1]C+1");
	vData.OlePropertyGet("Range").OlePropertyGet("Item", 2).OlePropertySet("FormulaR1C1", "=1");
	return this;
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

/* TExcelTableColumn* TExcelTableColumn::SetBorders() {}

TExcelTableColumn* TExcelTableColumn::SetWidth() {}
TExcelTableColumn* TExcelTableColumn::SetHeight() {}
TExcelTableColumn* TExcelTableColumn::AutoSize() {}

TExcelTableColumn* TExcelTableColumn::SetFormat() {} */

}

