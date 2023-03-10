#ifndef uTExcelObjectRangedTemplateH
#define uTExcelObjectRangedTemplateH

#include "uTExcelObjectTemplate.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Шаблон для классов, в которых есть диапазоны
 * @ingroup Templates
 * 
 * В дополнение к базовым методам содержит функционал для работы с 
 * диапазонами ячеек.
 * 
 * ---------------------------------------------------------------------- **/
template<class T>
class DLL_EI TExcelObjectRangedTemplate : public TExcelObjectTemplate<T>
{
public:
	TExcelObjectRangedTemplate(TExcelObject* pParent, const Variant& data);
	TExcelObjectRangedTemplate(const TExcelObjectRangedTemplate<T>& src);
//protected:
	~TExcelObjectRangedTemplate();

private:
    int ColStrToInteger(const String& str) const; ///< Преобразование буквы столбика в номер

protected:
    AnsiString ColToStrA(unsigned int ACol); ///< Преобразование номера в букву
    unsigned int GetColFromStr(const String& str); ///< Определить Колонку из диапазона
    unsigned int GetRowFromStr(const String& str); ///< Определить Строку

    /// Преобразовать выбор в строку
    AnsiString GetRangeString(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);
    /// Преобразовать позицию в строку
    AnsiString GetCellString(unsigned int col, unsigned int row);

    /// Проверка на корректность передачи столбика и колонки
    void checkColRow(unsigned int& col, unsigned int& row);

    /// Установить и выбрать дочернюю варианту на позицию
    void selectSingle(unsigned int col, unsigned int row);
    /// Установить и выбрать дочернюю варианту на диапазон
    void selectRange(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);
    /// Установить дочернюю варианту на позицию
    void setSingle(unsigned int col, unsigned int row);
    /// Установить дочернюю варианту на диапазон
    void setRange(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);

public:


};

}

//---------------------------------------------------------------------------
#endif

