//---------------------------------------------------------------------------

#ifndef uTExcelEnumsH
#define uTExcelEnumsH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
namespace exl {

// �������� ������ ��� ������� ������
enum class XlPivotTableSourceType : short int {
	xlConsolidation = 3,	// ��������� ���������� ������������
	xlDatabase = 1,			// Microsoft Excel ������ ��� ���� ������
	xlExternal = 2,			// ������ �� ������� ����������
	xlPivotTable = -4148	// ��� �� ��������, ��� � ������ ����� PivotTable
};

// ������� ��� ������� ������
enum class XlConsolidationFunction : short int { 
	xlAverage = -4106,		// ������� ��������
	xlCount = -4112,		// ����������
	xlCountNums = -4113,	// ���������� ������ ��������� ��������
	//xlDistinctCount = 11,	// ���������� ���������� �������� - �� �������� � ������ �������
	xlMax = -4136,			// ��������
	xlMin = -4139,			// �������
	xlProduct = -4149,		// ������������
	xlStDev = -4155,		// ����������� ����������, ���������� �� ���������� �������� (?)
	xlStDevP = -4156,		// ����������� ����������, ���������� �� ���� ���������
	xlSum = -4157,			// �����
	xlUnknown = 1000,		// ������� �� �������
	xlVar = -4164,			// �������, ���������� �� ���������� ��������
	xlVarP = -4165			// �������, ���������� �� ���� ���������
};

// ���������� ����� � ������� ��������
enum class XlPivotFieldOrientation : unsigned short int {
	xlColumnField = 2,	// �������
	xlDataField = 4,	// ������
	xlHidden = 0,		// ������
	xlPageField = 3,	// ���������
	xlRowField = 1		// ������
};

// ������������ ����� (�� ��� ���������, ���� ��� ���� �������...)
enum class ExcelTextAlign : short int {
	xlCenter = -4108,	// �����
	xlBottom = -4107,	// ���
	xlTop = -4160,		// ����
	xlLeft = -4131,		// ����
	xlRight = -4152		// �����
};

// ����������� ��� ���������� (?)
enum class FillDirection : char {
    Down = 0, Up, Left, Right
};

}
//---------------------------------------------------------------------------
#endif  // uTExcelEnumsH

