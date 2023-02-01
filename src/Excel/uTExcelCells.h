#ifndef uTExcelCellsH
#define uTExcelCellsH

#include "uTExcelObjectRangedTemplate.h"
#include "uIFormatManager.h"
//---------------------------------------------------------------------------
namespace exl {
/** @addtogroup ExcelClientObjects
 * @{
 * @brief Класс ячейки
 * 
 * Служит для взаимодействия с ячейкой или диапазоном ячеек. Предоставляет самые
 * востребованные методы для работы с ними.
 */
class DLL_EI TExcelCells : 
	public TExcelObjectRangedTemplate<TExcelCells>, 
	public IFormatManager<TExcelCells>
{
public:
	TExcelCells(TExcelObject* pParent, const Variant& data);
    TExcelCells(TExcelCells&);
    ~TExcelCells();

private:
	unsigned int dColumn, dRow;

public:
    unsigned int GetColumnsCount(); ///< Количество колонок
    unsigned int GetRowCount(); ///< Количество строк

	TExcelCells* GetCell(unsigned int col, unsigned int row);
	TExcelCells* GetCells(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);

    TExcelCells* Merge(); ///< Объеденить ячейки

	/// @{
	/// @brief Методы для вставки значений в ячейки.
	/// @warning При вставки значения в диапазон, значение будет вставлено в каждую ячейку
	TExcelCells* Insert(const Variant& data);
	TExcelCells* InsertString(const String& data);
	TExcelCells* InsertFormula(String& formula);
	/// @}

	/// @{
	/// @brief Вернуть значение ячейки
	/// @todo Узнать, что вернет из диапазона
    Variant ReadValue();
    String ReadValueString();
	/// @}

	// IFormatManager
	TExcelCells* SetHorizontalAlign(ExcelTextAlign align);
	TExcelCells* SetVerticalAlign(ExcelTextAlign align);

	/*
	TExcelCells* SetBorders();

	TExcelCells* SetWidth();
	TExcelCells* SetHeight();
	TExcelCells* AutoSize();

	TExcelCells* SetFormat();
	*/
};

}
/// @}
#endif

