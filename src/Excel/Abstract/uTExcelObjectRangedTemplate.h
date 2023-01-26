#ifndef uTExcelObjectRangedTemplateH
#define uTExcelObjectRangedTemplateH

#include "uTExcelObjectTemplate.h"
//---------------------------------------------------------------------------
namespace exl {
/** @addtogroup Templates
 * @{
 * 
 * @brief Шаблон для диапозонных элементов
 * 
 * В дополнение к базовым методам содержит функционал для работы с диапазонами
 * ячеек.
 */
template<class T>
class DLL_EI TExcelObjectRangedTemplate : public TExcelObjectTemplate<T>
{
public:
	TExcelObjectRangedTemplate(TExcelObject* pParent, const Variant& data);
	TExcelObjectRangedTemplate(const TExcelObjectRangedTemplate<T>& src);
protected:
	~TExcelObjectRangedTemplate();

private:
    int ColStrToInteger(const String& str) const; ///< Преобразование буквы столбика в номер

protected:
    AnsiString ColToStrA(unsigned int ACol); ///< Преобразование номера в букву
    unsigned int GetColFromStr(const String& str); ///< 
    unsigned int GetRowFromStr(const String& str);

    /// Преобразовать выбор в строку
    AnsiString GetRangeString(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);
    /// Преобразовать позицию в строку
    AnsiString GetCellString(unsigned int col, unsigned int row);

    /// Проверка на корректность передачи столбика и колонки
    void checkColRow(unsigned int& col, unsigned int& row);

    /// Установить дочернюю варианту на позицию
    void selectSingle(unsigned int col, unsigned int row);
    /// Установить дочернюю варианту на диапозон
    void selectRange(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);

public:


};

}

/// @}
#endif

