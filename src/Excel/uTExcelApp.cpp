//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelApp.h"
#include "uTExcelAppExceptions.h"

#include <fstream>

//#include "ExcelAPI.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
TExcelApp::TExcelApp()
	: TExcelObjectTemplate<TExcelApp>()
{
    Init();
}

TExcelApp::TExcelApp(bool visible)
	: TExcelObjectTemplate<TExcelApp>()
{
	Init();
	CreateApp(visible);
}

TExcelApp::TExcelApp(bool visible, unsigned int nSheetsInNewWorkbook)
	: TExcelObjectTemplate<TExcelApp>()
{
	Init();
	CreateApp(visible, nSheetsInNewWorkbook);
}

TExcelApp::TExcelApp(const TExcelApp& src)
	: TExcelObjectTemplate<TExcelApp>(src)
{
    Notifications = src.Notifications;
}

TExcelApp::~TExcelApp()
{
	using namespace std;

	ofstream log;
	log.open("ExcelUseLog.log", ios::app);

	if (log.is_open()) {
		String dtStamp = Now();

		log<<dtStamp.t_str()<<": ";
		log<<sizeof(*this)<<" bytes";
		log<<endl;

		log.close();
	}
}

void TExcelApp::Init() {
	Notifications = true;
	vData = Null();
}

TExcelApp* TExcelApp::CreateApp(bool visible) {
#ifdef EXCEL_APP_CREATED_ERROR
	if (!vData.IsNull()) throw ExcelAppAttachedException("CreateApp");
#endif

	if (vData.IsNull()){
		vData = Variant::CreateObject("Excel.Application");
		vData.OlePropertySet("Visible", visible);
	}
	return this;
}

TExcelApp* TExcelApp::CreateApp(bool visible, unsigned int nSheetsInNewWorkbook)
{
#ifdef EXCEL_APP_CREATED_ERROR
	if (!vData.IsNull()) throw ExcelAppAttachedException("CreateApp");
#endif

	if (nSheetsInNewWorkbook < 1) nSheetsInNewWorkbook = 1;
	if (vData.IsNull()) {
		vData = Variant::CreateObject("Excel.Application");
		vData.OlePropertySet("Visible", visible);
		vData.OlePropertySet("SheetsInNewWorkbook", nSheetsInNewWorkbook);
	}
	return this;
}

TExcelApp* TExcelApp::AttachApp() {
	if (!vData.IsNull()) throw ExcelAppAttachedException("AttachApp");

	vData = Variant::GetActiveObject("Excel.Application");
	return this;
}

TExcelApp* TExcelApp::TryAttachApp() {
	if (!vData.IsNull()) throw ExcelAppAttachedException("TryAttachApp");

	try {
		vData = Variant::GetActiveObject("Excel.Application");
	} catch (...) {
		vData = Variant::CreateObject("Excel.Application");
	}
	return this;
}

void TExcelApp::DeattachApp() {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("DeattachApp");
	//vData.Clear();
}

void TExcelApp::Close(bool silent) {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("Close");

	// Штатно закрыть эксель
    if (silent) vData.OlePropertySet("DisplayAlerts", false);
	
	vData.OleProcedure("Quit");
	vData.OlePropertySet("Interactive", true);
	vData.OlePropertySet("DisplayAlerts", true);
	vData = Null();
}

void TExcelApp::Free(){
	vData.OlePropertySet("DisplayAlerts", false);
	vData.OleProcedure("Quit");
	vData.OlePropertySet("Interactive", true);
	delete this;
}

TExcelApp* TExcelApp::SetExcelNotifications(bool stat){
	if (vData.IsNull()) throw ExcelAppNotAttachedException("SetExcelNotifications");

    Notifications = stat;
	vData.OlePropertySet("DisplayAlerts", stat);
	return this;
}

TExcelApp* TExcelApp::SetSheetsInNewWorkbook(unsigned int N) {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("SetSheetsInNewWorkbook");

	if (N < 1) N = 1;
	vData.OlePropertySet("SheetsInNewWorkbook", N);
	return this;
}

unsigned int TExcelApp::WorkbookCount() {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("WorkbookCount");

    return getChildCountByType("Workbooks");
}

TExcelWorkbook* TExcelApp::CreateWorkbook() {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("CreateWorkbook");

	vData.OlePropertyGet("Workbooks").OleFunction("Add");
	seekAndSetDataChild("Workbooks", getChildCountByType("Workbooks"));
    vDataChild.OleProcedure("Activate");	
	return GetCurrentWorkbook();
}

TExcelWorkbook* TExcelApp::CreateWorkbook(const String& workbookName) {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("CreateWorkbook");

    TExcelWorkbook* out = CreateWorkbook();
	if (out){
		//out->setName(workbookName);
	}
	return out;
}

TExcelWorkbook* TExcelApp::GetCurrentWorkbook() {
	if (vData.IsNull()) throw ExcelAppNotAttachedException("GetCurrentWorkbook");

	TExcelWorkbook* out = 0;
	if (vDataChild != Null()){
		out = new TExcelWorkbook(this, vDataChild);	
	}
	return out;
}

TExcelWorkbook* TExcelApp::OpenWorkbook(const String& path)
{
	if (vData.IsNull()) throw ExcelAppNotAttachedException("CreateWorkbook");

	vData.OlePropertyGet("Workbooks").OleProcedure("Open", System::StringToOleStr(path));
	seekAndSetDataChild("Workbooks", getChildCountByType("Workbooks"));

	TExcelWorkbook* out = new TExcelWorkbook(this, vDataChild);
	return out;
}

}


