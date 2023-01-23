//---------------------------------------------------------------------------

#ifndef uTExcelObjectRangedTemplateH
#define uTExcelObjectRangedTemplateH

//---------------------------------------------------------------------------
#include "uTExcelObjectTemplate.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
template<class T>
class DLL_EI TExcelObjectRangedTemplate : public TExcelObjectTemplate<T>
{
public:
	TExcelObjectRangedTemplate(TExcelObject* pParent, const Variant& data);
	TExcelObjectRangedTemplate(const TExcelObjectRangedTemplate<T>& src);
protected:
	~TExcelObjectRangedTemplate();

private:
    int ColStrToInteger(const String& str) const;

protected:
    AnsiString ColToStrA(unsigned int ACol);
    unsigned int GetColFromStr(const String& str);
    unsigned int GetRowFromStr(const String& str);

    AnsiString GetRangeString(
        unsigned int startColumn, unsigned int startRow,
        unsigned int endColumn, unsigned int endRow
            );
    AnsiString GetCellString(unsigned int col, unsigned int row);

    void checkColRow(unsigned int& col, unsigned int& row);

    //TExcelObject* selectChild(const String& childName);

    void selectSingle(unsigned int col, unsigned int row);
    void selectRange(
		unsigned int startColumn, unsigned int startRow,
        unsigned int endColumn, unsigned int endRow
	);

public:


};

}

//---------------------------------------------------------------------------
#endif

