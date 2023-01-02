//---------------------------------------------------------------------------

#ifndef uTExcelTableHeadersH
#define uTExcelTableHeadersH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../uTExcelCells.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelTableHeaders : public TExcelObjectRanged
{
public:
	TExcelTableHeaders(TExcelObjectRanged* pParent, const Variant& data);
    TExcelTableHeaders(TExcelObjectRanged* Source)
        : TExcelObjectRanged(Source->getParent(), Source->getVariant())
        {}
        
    ~TExcelTableHeaders();

public:


public:
    //unsigned int HeadersDepth();
    //unsigned int HeadersCount();
    //
    //TExcelCell* GetHeader(unsigned int col);
    //TExcelCell* GetHeader(unsigned int col, unsigned int depth);

};
//---------------------------------------------------------------------------
}	// end namespace exl
//---------------------------------------------------------------------------
#endif

