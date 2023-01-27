#ifndef uTExcelTableColumnH
#define uTExcelTableColumnH

#include "uDll.h"
#include "uTExcelCells.h"
//---------------------------------------------------------------------------
namespace exl {
/** @addtogroup Tables
 * @{
 * 
 * @brief Класс столбика таблицы
 * 
 * Если честно, то хз зачем
 * 
 */
class DLL_EI TExcelTableColumn : public TExcelObjectRangedTemplate<TExcelTableColumn>
                               //, public IFormatManager<TExcelTableColumn>
{
public:
	TExcelTableColumn(TExcelObject* pParent, const Variant& data);
    TExcelTableColumn(TExcelTableColumn& src);
    ~TExcelTableColumn();

public:
	TExcelTableColumn* SetIdentity(int start, int step); ///< Установить колонку как счетчик

	TExcelTableColumn* SetHorizontalAlign(ExcelTextAlign align);
	TExcelTableColumn* SetVerticalAlign(ExcelTextAlign align);

	TExcelTableColumn* SetBorders();

	TExcelTableColumn* SetWidth();
	TExcelTableColumn* SetHeight();
	TExcelTableColumn* AutoSize();

	TExcelTableColumn* SetFormat();

};

}
//---------------------------------------------------------------------------
#endif // uTExcelTableColumnH

