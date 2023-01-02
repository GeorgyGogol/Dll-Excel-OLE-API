//---------------------------------------------------------------------------

#ifndef uTPivotSettingsH
#define uTPivotSettingsH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../Table/uTTableCreator.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
// ���������� (���������) ������ � ������� �������
enum PivotDataPlace : char { Rows = 0, Columns = 1 };

//---------------------------------------------------------------------------
// ��������� - ������� � ��� ���������
struct DLL_EI TExcelTablePivotField
{
	String ColumnName;
	String Caption;

	XlPivotFieldOrientation Type;
	XlConsolidationFunction Function;
};

//---------------------------------------------------------------------------
// �����-��������� ��� ��������� ������� �������
class DLL_EI TPivotSettings
{
public:
	TPivotSettings(const String& pivotName);
	~TPivotSettings();

	typedef std::vector<TExcelTablePivotField> tSettingsArray;

private:
	String PivotName;		// �������� �������

public:
	String NewSheetName;
	
	String GetPivotName();
	void SetPivotName(const String& name);

	String PivotTitle;		// ���������
	tSettingsArray Settings;	// ������ ��������, �� � ����� �������
	PivotDataPlace DataPlace; // ������������ �������� - �� �������� ��� ���������
	bool ShowRowTotal;		// �������� ����� �� ��������
	bool ShowColumnTotal;	// �������� ����� �� �������

	void Clear(); // �������� ���������
	void Add(const TExcelTablePivotField& setting);	// �������� 
	void AddRow(const String& �olumnName, const String& caption = "");
	void AddColumn(const String& �olumnName, const String& caption = "");
	void AddData(const String& �olumnName, const String& caption = "",
		XlConsolidationFunction function = XlConsolidationFunction::xlUnknown);

	unsigned int RowCount() const;
	unsigned int ColumnCount() const;
};
//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

