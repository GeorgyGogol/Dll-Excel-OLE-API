//---------------------------------------------------------------------------

#pragma hdrstop

#include "uTExcelCell.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {

/* TExcelCell::TExcelCell() 
	: TExcelObjectRanged()
{
} */

TExcelCell::TExcelCell(TExcelObjectRanged* pParent, const Variant& data)
    : TExcelObjectRanged(pParent, data)
{
}

TExcelCell::TExcelCell(const TExcelCell& src) 
	: TExcelObjectRanged(src)
{
}

TExcelCell::~TExcelCell() {
}

/* TExcelObject* TExcelCell::getCurrentSelectedChild() {
	TExcelObject* out = new TExcelCell(this, vDataChild);
	Childs.push_back(out);
	return out;
} */

TExcelCell* TExcelCell::Insert(const Variant& data) {
    vData.OlePropertySet("Value", data);
    return this;
}

TExcelCell* TExcelCell::InsertString(const String& data) {
    vData.OlePropertySet("Value", System::StringToOleStr(data));
    return this;
}

TExcelCell* TExcelCell::InsertFormula(String& formula) {
    if (formula.SubString(0, 1) != "=") formula = "=" + formula;

    vData.OlePropertySet("Formula", System::StringToOleStr(formula));

    return this;
}

TExcelCell* TExcelCell::Merge() {
	// Для начала нам нужно выключить предупреждение. Получим Эксель
	Variant vApp = ((TExcelObject*)(getParent()->getParent()))->getParentVariant();
	bool notifications = vApp.OlePropertyGet("DisplayAlerts"); // Запомним состояние Алертов
	vApp.OlePropertySet("DisplayAlerts", false);	// Выключим

	// Здесь сделаем наше грязное дело, совсем как в Экселе
	vData.OleProcedure("Merge");
	SetHorizontalAlign(ExcelTextAlign::xlCenter);
	SetVerticalAlign(ExcelTextAlign::xlCenter);

	// Вернем предупреждения
	vApp.OlePropertySet("DisplayAlerts", notifications);

    return this;
}

TExcelCell* TExcelCell::SetHorizontalAlign(ExcelTextAlign align) {
    vData.OlePropertySet("HorizontalAlignment", align);
    return this;
}

TExcelCell* TExcelCell::SetVerticalAlign(ExcelTextAlign align) {
    vData.OlePropertySet("VerticalAlignment", align);
    return this;
}

String TExcelCell::ReadString() {
	String out = VarToStr(vData.OlePropertyGet("Value"));
	return out;
}

}


