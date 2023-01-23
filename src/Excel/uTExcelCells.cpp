//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelCells.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelCells::TExcelCells(TExcelObject* pParent, const Variant& data) 
	: TExcelObjectRangedTemplate<TExcelCells>(pParent, data) //, IFormatManager<TExcelCells>()
{
	//String AddressString = VarToStr(vData.OlePropertyGet("Address"));
}

TExcelCells::TExcelCells(TExcelCells& src) 
	: TExcelObjectRangedTemplate<TExcelCells>(src) //, IFormatManager<TExcelCells>()
{
}

TExcelCells::~TExcelCells() 
{}

unsigned int TExcelCells::GetColumnsCount() 
{
    return getChildCountByType("Columns");
}

unsigned int TExcelCells::GetRowCount() 
{
    return getChildCountByType("Rows");
}

TExcelCells* TExcelCells::GetCell(unsigned int col, unsigned int row)
{
	selectSingle(col, row);
	TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

TExcelCells* TExcelCells::GetCells(
    unsigned int startColumn, unsigned int startRow,
    unsigned int endColumn, unsigned int endRow
) 
{
	selectRange(startColumn, startRow, endColumn, endRow);
	TExcelCells* out = new TExcelCells(this, vDataChild);
    return out;
}

TExcelCells* TExcelCells::Merge() {
    checkDataValide();
    
	// Для начала нам нужно выключить предупреждение. Найдем Эксель (Самый галвный тут)
	TExcelObject* pApp = GetParent();
	while (pApp->GetParent()){
		pApp = pApp->GetParent();
	}

	Variant vApp = pApp->getVariant(); // Возьмем варианту Экселя
	Variant notifications = vApp.OlePropertyGet("DisplayAlerts"); // Запомним состояние Алертов
	vApp.OlePropertySet("DisplayAlerts", false);	// Выключим

	// Здесь сделаем наше грязное дело, совсем как в Экселе
	vData.OleProcedure("Merge");
	SetHorizontalAlign(ExcelTextAlign::xlCenter);
	SetVerticalAlign(ExcelTextAlign::xlCenter);

	// Вернем предупреждения
	vApp.OlePropertySet("DisplayAlerts", notifications);

    return this;
}


TExcelCells* TExcelCells::Insert(const Variant& data)
{
    checkDataValide();
    vData.OlePropertySet("Value", data);
    return this;
}

TExcelCells* TExcelCells::InsertString(const String& data)
{
    checkDataValide();
    vData.OlePropertySet("Value", System::StringToOleStr(data));
    return this;
}

TExcelCells* TExcelCells::InsertFormula(String& formula)
{
    checkDataValide();
    if (formula.SubString(0, 1) != "=") formula = "=" + formula;
    vData.OlePropertySet("Formula", System::StringToOleStr(formula));
    return this;
}

Variant TExcelCells::ReadValue(){
	checkDataValide();
	return vData.OlePropertyGet("Value");
}

String TExcelCells::ReadValueString() {
	checkDataValide();
	String out = VarToStr(vData.OlePropertyGet("Value"));
	return out;
}

TExcelCells* TExcelCells::SetHorizontalAlign(ExcelTextAlign align)
{
    checkDataValide();
	vData.OlePropertySet("HorizontalAlignment", (int)align);
    return this;
}

TExcelCells* TExcelCells::SetVerticalAlign(ExcelTextAlign align)
{
    checkDataValide();
	vData.OlePropertySet("VerticalAlignment", (int)align);
    return this;
}

TExcelCells* TExcelCells::SetBorders()
{
    checkDataValide();
    // TODO
    return this;
}

TExcelCells* TExcelCells::SetWidth()
{
    checkDataValide();
    // TODO
    return this;
}

TExcelCells* TExcelCells::SetHeight()
{
    checkDataValide();
    // TODO
    return this;
}

TExcelCells* TExcelCells::AutoSize()
{
    checkDataValide();
    // TODO
    return this;
}

TExcelCells* TExcelCells::SetFormat()
{
    checkDataValide();
    // TODO
    return this;
}

}


