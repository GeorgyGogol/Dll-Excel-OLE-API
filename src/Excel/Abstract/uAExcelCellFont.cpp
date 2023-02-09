//---------------------------------------------------------------------------


#pragma hdrstop

#include "uAExcelCellFont.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------

/// @details Для создания экземпляра нам просто необходим класс
/// @param src 
ExcelCellFont::ExcelCellFont(TExcelObject* src)
{
    // TODO: Проверка на существование
    pVarCell = &(src->GetVariant());
}

ExcelCellFont::~ExcelCellFont()
{
    pVarCell = NULL;
}

ExcelCellFont ExcelCellFont::setBold(bool isBold)
{
    pVarCell->OlePropertySet("Bold", isBold);
}

ExcelCellFont ExcelCellFont::setItalic(bool isItalic)
{
    pVarCell->OlePropertySet("Italic", isItalic);
}

/* 
ExcelCellFont ExcelCellFont::setFontSize(unsigned int points)
{

}

ExcelCellFont ExcelCellFont::setFontName(const String& name)
{

}
 */

/// @param points Пункты ширины
ExcelCellFont ExcelCellFont::setRowHeight(unsigned int points)
{
    pVarCell->OlePropertySet("RowHeight", points);
}

/// @details Метод позволяет установить высоту строки без вычислений, сколько
/// пунктов для этого нужно.
/// @param lines Количество строк
/// @todo Вычисление размера шрифта
ExcelCellFont ExcelCellFont::setRowHeightInTextLines(unsigned int lines)
{
    pVarCell->OlePropertySet("RowHeight", lines * 15);
}

/* 
ExcelCellFont ExcelCellFont::setColumnWidth(unsigned int points)
{

}
 */

ExcelCellFont ExcelCellFont::ColumnAutoFit()
{
    pVarCell->OlePropertyGet("EntireColumn").OleProcedure("AutoFit");
}


}

