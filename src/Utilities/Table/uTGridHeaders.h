#ifndef uTGridHeadersH
#define uTGridHeadersH

#include "uTExcelTable.h"
//---------------------------------------------------------------------------
namespace exl {

struct DLL_EI TGridHeader {
	String Caption;
	String FieldName;
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
	void Add(const TGridHeader& header);
	unsigned int CountVisible() const;
	unsigned int Count() const;
	unsigned int Deep() const;

    iterator Begin();
    iterator End();
    iterator Next();
	iterator NextVisible();
	bool Eof();

	TGridHeader* CurrentHeader();
	TGridHeader* GetHeader(unsigned int N);

	Variant generateVariant();

};

}

//---------------------------------------------------------------------------
#endif


