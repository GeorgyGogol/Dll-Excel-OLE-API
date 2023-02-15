#ifndef uTExcelCellsH
#define uTExcelCellsH

#include "uDll.h"
#include "uTExcelObjectRangedTemplate.h"
#include "uIFormatManager.h"
#include <uIBorderManager.h>
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс диапазонов ячеек
 * @ingroup ExcelClientObjects
 * 
 * Служит для взаимодействия с ячейкой или диапазоном ячеек. Предоставляет
 * самые востребованные методы для работы с ними.
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelCells : 
	public TExcelObjectRangedTemplate<TExcelCells>, 
	public IFormatManager<TExcelCells>,
	public IBorderManager<TExcelCells>
{
public:
	TExcelCells(TExcelObject* pParent, const Variant& data);
    TExcelCells(TExcelCells&);
    ~TExcelCells();

private:
	//unsigned int dColumn, dRow;

public:
    unsigned int GetColumnsCount(); ///< Количество колонок
	unsigned int GetRowCount(); ///< Количество строк

	void Delete();

	/// @brief Получить ячейку из диапазона
	TExcelCells* GetCell(unsigned int col, unsigned int row);

	/// @brief Получить "поддиапазон" из диапазона
	TExcelCells* GetCells(unsigned int startColumn, unsigned int startRow, unsigned int endColumn, unsigned int endRow);

	/// @brief Методы для вставки значений в ячейки.
	/// @warning При вставки значения в диапазон, значение будет вставлено в каждую ячейку
	TExcelCells* Insert(const Variant& data);
	TExcelCells* InsertString(const String& data);
	TExcelCells* InsertFormula(String& formula);

	/// @brief Вернуть значение ячейки
	/// @todo Узнать, что вернет из диапазона
    Variant ReadValue();
    String ReadValueString();

	TExcelCells* Merge();

	// IFormatManager override
	IFormatManager<TExcelCells>* GetFormatInterface();
	TExcelCells* SetHorizontalAlign(ExcelTextAlign align);
	TExcelCells* SetVerticalAlign(ExcelTextAlign align);
	TExcelCells* SetTextWrap(bool state);
	TExcelCells* SetFormat(ExcelFormats format);
	TExcelCells* SetFormat(const String& format);

	// IBorderManager override
	IBorderManager<TExcelCells>* GetBorderInterface();
	//TExcelCells* SetBordersAll(XlLineStyle style, XlBorderWeight weight);
	TExcelCells* SetBorder(XlBordersIndex border, XlLineStyle style, XlBorderWeight weight);
	//TExcelCells* RemoveBordersAll();
	TExcelCells* RemoveBorder(XlBordersIndex border);
	
};

}

//---------------------------------------------------------------------------
#endif

