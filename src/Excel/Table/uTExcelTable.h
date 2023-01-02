//---------------------------------------------------------------------------

#ifndef uTExcelTableH
#define uTExcelTableH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelTableHeaders.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelTable : public TExcelObject
{
public:
	//TExcelTable(TExcelObjectRanged* sheet, const Variant& name, const Variant& headers, const Variant& data);
	TExcelTable(TExcelObjectRanged* sheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data, TExcelCell* titleCells);
	TExcelTable(TExcelObjectRanged* sheet, const String& tableName, TExcelTableHeaders* headers, TExcelCells* data);
    //TExcelTable(TExcelObjectRanged* sheet, TExcelCell* cells, unsigned int headersCount);
    ~TExcelTable();

private:
    String TableName;
	TExcelTableHeaders* Headers;
	TExcelCells* Data;
    TExcelCell* TitleCells;

public:
    String GetTitle();
    TExcelTable* SetTitle(const String& title);

    TExcelTableHeaders* GetHeaders();
    //TExcelCell* GetCell(unsigned int col, unsigned int row);
};

//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif


