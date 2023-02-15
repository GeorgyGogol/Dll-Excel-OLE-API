#ifndef uTExcelTableColumnH
#define uTExcelTableColumnH

#include "uDll.h"
#include "uTExcelCells.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс столбца таблицы
 * @ingroup Tables
 * 
 * Нет, столбик таблицы не является прям отражением диапазона ячеек. Да, это
 * диапазон, но есть нюанс:
 * 1. В целях безопасности я убрал методы вставки значений, т.к. при такой 
 * операции будут заменены ВСЕ значения в столбце
 * 2. Этот объект имеет некторые приколюхи, не доступные для обычного 
 * диапазона.
 * 
 * В итоге мы получаем несколько другую трактовку ячеек.
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelTableColumn 
	: public TExcelObjectRangedTemplate<TExcelTableColumn>
	, public IFormatManager<TExcelTableColumn>
{
public:
	TExcelTableColumn(TExcelObject* pParent, const Variant& data);
    TExcelTableColumn(TExcelTableColumn& src);
    ~TExcelTableColumn();

public:
	//TExcelTableColumn* SetIdentity(int start, int step); ///< Установить колонку как счетчик

	// IFormatManager override
	IFormatManager<TExcelTableColumn>* GetFormatInterface();
	TExcelTableColumn* SetHorizontalAlign(ExcelTextAlign align);
	TExcelTableColumn* SetVerticalAlign(ExcelTextAlign align);
	TExcelTableColumn* SetTextWrap(bool state);
	TExcelTableColumn* SetFormat(ExcelFormats format);
	TExcelTableColumn* SetFormat(const String& format);

};

}

//---------------------------------------------------------------------------
#endif // uTExcelTableColumnH

