#ifndef uAExcelCellFontH
#define uAExcelCellFontH
/** @file
 * @brief �������� ���������� ���������� �������
 */

#include "uTExcelObject.h"
//---------------------------------------------------------------------------
namespace exl{
/** -------------------------------------------------------------------------
 * @brief ��������� ��� ��������� ������� �����
 * @ingroup Abstract
 * 
 * �������������� � �������� ������ ��� ������ �� ��������.
 * 
 * � �������� ������ ����������, ������������ ��� - ��� ������ �����
 * ��������� �������� (����� �� �������� ������ ������ �������)
 * 
 * @todo �������� �������������� ������, � �� ������ ��, �������
 * ��������������
 * 
 * ---------------------------------------------------------------------- **/
class ExcelCellFont
{
private:
    Variant* pVarCell; ///< ��������� �� �������� �������-���������

public:
    ExcelCellFont(TExcelObject* src); ///< ������� �����������
    ~ExcelCellFont();

    ExcelCellFont setBold(bool isBold); ///< ���������� ��������
    ExcelCellFont setItalic(bool isItalic); ///< ���������� ������
    //ExcelCellFont setBold();

    /*
    ExcelCellFont setFontSize(unsigned int points);
    ExcelCellFont setFontName(const String& name);
    */

    /// @brief ���������� ������ ����� � �������
    ExcelCellFont setRowHeight(unsigned int points);
    /// @brief ������ ����� � �������� ������
    ExcelCellFont setRowHeightInTextLines(unsigned int lines);

    /*
    ExcelCellFont setColumnWidth(unsigned int points);
    */

    ExcelCellFont ColumnAutoFit(); ///< ���������� �������

};

}

//---------------------------------------------------------------------------
#endif
