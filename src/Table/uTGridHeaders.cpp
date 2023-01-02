//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTGridHeaders.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {

TGridHeaders::TGridHeaders(TDataSet* dataSet)
	: nVisible(0)
{
	TGridHeader buf;
	buf.Visible = true;
	for (int tabCol = 0; tabCol < dataSet->Fields->Count; tabCol++) {
		buf.Caption = dataSet->Fields->Fields[tabCol]->FieldName;
		Add(buf);
	}
}

TGridHeaders::TGridHeaders(TDBGridEh* gridEh)
	: nVisible(0)
{
	TGridHeader buf;
	for (int tabCol = 0; tabCol < gridEh->Columns->Count; tabCol++) {
		buf.Caption = gridEh->Columns->Items[tabCol]->Title->Caption;
		buf.Visible = gridEh->Columns->Items[tabCol]->Visible;

		Add(buf);
	}
    // (�.�������, 08.09.2022 14:40) TODO: ����������� ��������� ���������
    // �������������� 22.12.2022. ��� ��� ����� ������, � �������� ��� � ��
    // ������.
}

void TGridHeaders::Add(TGridHeader header) {
	if (header.Visible) nVisible++;
	Headers.push_back(header);

	it = Headers.end();
}

unsigned int TGridHeaders::CountVisible() const
{
    return nVisible;
}

unsigned int TGridHeaders::Count() const
{
	return Headers.size();
}

unsigned int TGridHeaders::Deep() const {
    return 1;
}

TGridHeaders::iterator TGridHeaders::Begin() {
	it = Headers.begin();
    return it;
}

TGridHeaders::iterator TGridHeaders::End() {
	it = Headers.end();
    return it;
}

TGridHeaders::iterator TGridHeaders::Next() {
	if (it == Headers.end()) it = Headers.end();
	return it;
}

TGridHeaders::iterator TGridHeaders::NextVisible() {
	++it;
	while (!it->Visible && it != Headers.end())
    {
        ++it;
    }
	return it;
}

bool TGridHeaders::Eof() {
	return it == Headers.end();
}

TGridHeader TGridHeaders::CurrentHeader() {
    return (*it);
}

TGridHeader TGridHeaders::GetHeader(unsigned int N) {
    return Headers[N];
}

Variant TGridHeaders::generateVariant() {
	Variant varHeaders(OPENARRAY(int, (1, 1, 1, nVisible)), varString); // !!!
	
	for (unsigned int tabCol = 1; !Eof(); NextVisible(), ++tabCol)
	{
		varHeaders.PutElement(CurrentHeader().Caption, 1, tabCol); // 1 - ������
    }
	
	return varHeaders;
}

}

