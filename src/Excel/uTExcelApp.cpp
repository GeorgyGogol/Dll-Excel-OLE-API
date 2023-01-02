//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelApp.h"
#include "uTExcelWorkbook.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
TExcelApp::TExcelApp()
	: TExcelObject()
{
    Init();
}

TExcelApp::TExcelApp(bool visible)
	: TExcelObject()
{
	Init();
	CreateApp(visible);
}

TExcelApp::TExcelApp(bool visible, unsigned int nSheetsInNewWorkbook)
	: TExcelObject()
{
	Init();
	CreateApp(visible, nSheetsInNewWorkbook);
}

TExcelApp::TExcelApp(const TExcelApp& src)
	: TExcelObject(src)
{
    Notifications = src.Notifications;
}

TExcelApp::~TExcelApp()
{}

void TExcelApp::Init() {
    Notifications = true;
}

TExcelApp* TExcelApp::TryAttachApp() {
	try {
		vData = Variant::GetActiveObject("Excel.Application");
	} catch (...) {
		vData = Variant::CreateObject("Excel.Application");
	}
	return this;
}

TExcelApp* TExcelApp::CreateApp(bool visible) {
	if (vData.IsNull()){
		vData = Variant::CreateObject("Excel.Application");
		vData.OlePropertySet("Visible", visible);
	}
	return this;
}

TExcelApp* TExcelApp::CreateApp(bool visible, unsigned int nSheetsInNewWorkbook)
{
	if (nSheetsInNewWorkbook < 1) nSheetsInNewWorkbook = 1;
	if (vData.IsNull()){
		vData = Variant::CreateObject("Excel.Application");
		vData.OlePropertySet("Visible", visible);
		vData.OlePropertySet("SheetsInNewWorkbook", nSheetsInNewWorkbook);
	}
	return this;
}

TExcelApp* TExcelApp::AttachApp() {
	vData = Variant::GetActiveObject("Excel.Application");
	return this;
}

void TExcelApp::DeattachApp() {
	//vData.Clear();
}

void TExcelApp::Close(bool silent) {
	// Штатно закрыть эксель
    if (silent) vData.OlePropertySet("DisplayAlerts", false);
	
	vData.OleProcedure("Quit");
	vData.OlePropertySet("Interactive", true);
	vData.OlePropertySet("DisplayAlerts", true);
	//vData = Null();
}

TExcelApp* TExcelApp::SetExcelNotifications(bool stat){
    Notifications = stat;
	vData.OlePropertySet("DisplayAlerts", stat);
	return this;
}

TExcelApp* TExcelApp::SetSheetsInNewWorkbook(unsigned int N) {
	if (N < 1) N = 1;
	vData.OlePropertySet("SheetsInNewWorkbook", N);
	return this;
}

unsigned int TExcelApp::WorkbookCount() {
    return getChildCountByType("Workbooks");
}

TExcelWorkbook* TExcelApp::CreateWorkbook() {
	vData.OlePropertyGet("Workbooks").OleFunction("Add");
	seekAndSetDataChild("Workbooks", getChildCountByType("Workbooks"));
    vDataChild.OleProcedure("Activate");	
	return GetCurrentWorkbook();
}

TExcelWorkbook* TExcelApp::CreateWorkbook(const String& workbookName) {
    TExcelWorkbook* out = CreateWorkbook();
	if (out){
		//out->setName(workbookName);
	}
	return out;
}

TExcelWorkbook* TExcelApp::GetCurrentWorkbook() {
	TExcelWorkbook* out = 0;
	if (vDataChild != Null()){
		out = new TExcelWorkbook(this, vDataChild);	
	}
	return out;
}


}

