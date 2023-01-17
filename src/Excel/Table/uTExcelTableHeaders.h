//---------------------------------------------------------------------------

#ifndef uTExcelTableHeadersH
#define uTExcelTableHeadersH

//---------------------------------------------------------------------------
// Copyright (c) 2022-2023 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../uTExcelCells.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelTableHeaders : private TExcelCells
{
public:
	TExcelTableHeaders(TExcelObject* pParent, const Variant& data);
    TExcelTableHeaders(TExcelObject* Source);
        
    ~TExcelTableHeaders();

public:


public:
	unsigned int HeadersDepth();
    unsigned int HeadersCount();

    String GetTitle(unsigned int i);
    void SetTitle(unsigned int i, const String& title);

    //
    //TExcelCell* GetHeader(unsigned int col);
    //TExcelCell* GetHeader(unsigned int col, unsigned int depth);

};
//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif

