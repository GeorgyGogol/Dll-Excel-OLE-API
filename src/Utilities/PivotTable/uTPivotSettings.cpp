//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTPivotSettings.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {

TPivotSettings::TPivotSettings()
{
	Clear();
}

TPivotSettings::TPivotSettings(const String& pivotName)
	: PivotName(pivotName)
{
	Clear();
}

TPivotSettings::~TPivotSettings() {}

String TPivotSettings::GetPivotName() {
	return PivotName;
}

void TPivotSettings::SetPivotName(const String& name) {
	if (name.Length() < 1) return;
	PivotName = name;
}

void TPivotSettings::Clear(){
	NewSheetName = "";
	PivotTitle = "";
    Settings.clear();
	DataPlace = PivotDataPlace::Rows;
	ShowRowTotal = true;
	ShowColumnTotal = true;
}

void TPivotSettings::Add(const TExcelTablePivotField& setting) {
	Settings.push_back(setting);
}

void TPivotSettings::AddRow(const String& name, const String& caption){
	TExcelTablePivotField add;
	add.ColumnName = name;
	add.Caption = caption.Length() > 0 ? caption : name;
	add.Type = XlPivotFieldOrientation::xlRowField;
	add.Function = XlConsolidationFunction::xlUnknown;

	//Rows.push_back(add);
	Settings.push_back(add);
}

void TPivotSettings::AddColumn(const String& name, const String& caption){
	TExcelTablePivotField add;
	add.ColumnName = name;
	add.Caption = caption.Length() > 0 ? caption : name;
	add.Type = XlPivotFieldOrientation::xlColumnField;
	add.Function = XlConsolidationFunction::xlUnknown;

	//Columns.push_back(add);
    Settings.push_back(add);
}

void TPivotSettings::AddData(const String& name, const String& caption, XlConsolidationFunction function){
	TExcelTablePivotField add;
	add.ColumnName = name;
	add.Caption = caption.Length() > 0 ? caption : name;
	add.Type = XlPivotFieldOrientation::xlDataField;
	add.Function = function;

	//Data.push_back(add);
    Settings.push_back(add);
}

unsigned int TPivotSettings::RowCount() const {
	unsigned int out = 0;
	for (unsigned int i = 0; i < Settings.size(); ++i){
		if (Settings[i].Type == XlPivotFieldOrientation::xlRowField) out++;
	}
	return out;
}

unsigned int TPivotSettings::ColumnCount() const {
	unsigned int out = 0;
	for (unsigned int i = 0; i < Settings.size(); ++i){
		if (Settings[i].Type == XlPivotFieldOrientation::xlColumnField) out++;
	}
	return out;
}

}

