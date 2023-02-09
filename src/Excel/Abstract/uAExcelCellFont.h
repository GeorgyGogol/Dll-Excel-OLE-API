#ifndef UAEXCELCELLFONT_H
#define UAEXCELCELLFONT_H
/** @file
 * @brief Содержит абстракцию управления шрифтом
 */

#include "uTExcelObject.h"
//---------------------------------------------------------------------------
namespace exl{
/** -------------------------------------------------------------------------
 * @brief Интерфейс для настройки формата ячеек
 * @ingroup Abstract
 * 
 * Приспособленец с методами только для работы со шрифтами.
 * 
 * В основном методы тривиальны, единственное что - это работа через
 * указатель варианты (чтобы не забивать память копией объекта)
 * 
 * @todo Добавить дополнительные методы, а не только те, которые 
 * использовались
 * 
 * ---------------------------------------------------------------------- **/
class ExcelCellFont
{
private:
    Variant* pVarCell; ///< Указатель на варианту объекта-оригинала
    
public:
    ExcelCellFont(TExcelObject* src); ///< Базовый конструктор
    ~ExcelCellFont();

    ExcelCellFont setBold(bool isBold); ///< Установить жирность
    ExcelCellFont setItalic(bool isItalic); ///< Установить курсив
    //ExcelCellFont setBold();

    /* 
    ExcelCellFont setFontSize(unsigned int points);
    ExcelCellFont setFontName(const String& name);
    */

    /// @brief Установить высоту строк в пунктах
    ExcelCellFont setRowHeight(unsigned int points); 
    /// @brief Выстоа строк в размерах шрифта
    ExcelCellFont setRowHeightInTextLines(unsigned int lines);

    /*
    ExcelCellFont setColumnWidth(unsigned int points);
    */

    ExcelCellFont ColumnAutoFit(); ///< Автоширина колонок

};

}

//---------------------------------------------------------------------------
#endif

