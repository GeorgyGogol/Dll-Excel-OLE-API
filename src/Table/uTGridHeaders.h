//---------------------------------------------------------------------------

#ifndef uTGridHeadersH
#define uTGridHeadersH

//---------------------------------------------------------------------------
#include "..\Excel\Table\uTExcelTable.h"
#include <vector.h>
#include <DB.hpp>
#include <ADODB.hpp>
#include "DBGridEh.hpp"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
struct DLL_EI TGridHeader {
	String Caption;
	bool Visible;
};
//---------------------------------------------------------------------------
class DLL_EI TGridHeaders
{
public:
	TGridHeaders(TDataSet* gridEh);
	TGridHeaders(TDBGridEh* gridEh);
	//~TGridHeaders();

	typedef std::vector<TGridHeader>::iterator iterator;

private:
	std::vector<TGridHeader> Headers;
    iterator it;
    unsigned int nVisible;

public:
	void Add(TGridHeader header);
	unsigned int CountVisible() const;
	unsigned int Count() const;
	unsigned int Deep() const;

    iterator Begin();
    iterator End();
    iterator Next();
	iterator NextVisible();
	bool Eof();

	TGridHeader CurrentHeader();
	TGridHeader GetHeader(unsigned int N);

	Variant generateVariant();

	/*void TTableCreator::generateHeadersVariant() {
	if (!Headers) throw ExcelTableCreatorException("generateHeadersVariant", 
		"Cannot generate Headers data because Headers not exists!");


	Variant varArr(OPENARRAY(int, (1, ArrSize, 1, nColumns)), varString); // !!!!!
   
    // Проработаем заголовки
	TExcelTableHeaders* tableHeaders;
	Variant varHeaders(OPENARRAY(int, (1, Headers->Count(), 1, Headers->CountVisible())), varVariant);
	// Занесем данные в массив
	//	unsigned int tabCol = 1; // Счетчик (чет не получилось в цикл запихнуть)
	Headers->Begin();
	for (unsigned int tabCol = 1; !Headers->Eof(); Headers->NextVisible(), ++tabCol)
	{
		varHeaders.PutElement(Headers->CurrentHeader().Caption, 1, tabCol); // 1 - строка
    }
	

	varHeaders = varArr;
}*/

};
//---------------------------------------------------------------------------
} 
//---------------------------------------------------------------------------
#endif
