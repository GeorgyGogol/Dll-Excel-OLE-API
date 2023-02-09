//---------------------------------------------------------------------------
#pragma hdrstop

#include "uTExcelApp.h"
#include "uTExcelAppExceptions.h"
#include "Log.h"

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

TExcelApp::TExcelApp(const TExcelApp& src)
	: TExcelObjectTemplate<TExcelApp>(src)
{
    Notifications = src.Notifications;
}

TExcelApp::~TExcelApp()
{
#ifdef ENABLE_USAGE_STATISTIC
	TLog log(CallLogReason::WriteUsedMemory);
	log.WriteSize(SizeOfThis());
#endif
}

/** @ingroup ExcelClientObjects
 * @defgroup ExcelInitMethods Методы подключения к Excel
 * 
 * Методы, с помощью которых устанавливается связь с приложением.
 * 
 * @param visible Видимость
 * @param nSheetsInNewWorkbook Количество листов в новой книге
 * @{
 */

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
/// @}

/// @details Иницилизация при создании объекта
void TExcelApp::Init() {
	Notifications = true;
	vData = Null();
}

/// @ingroup ExcelInitMethods
/// @{
TExcelApp* TExcelApp::CreateApp(bool visible) {
	if (!vData.IsNull()) throw ExcelAppAttachedException("CreateApp");

	if (vData.IsNull()){
		vData = Variant::CreateObject("Excel.Application");
		vData.OlePropertySet("Visible", visible);
	}
	return this;
}

TExcelApp* TExcelApp::CreateApp(bool visible, unsigned int nSheetsInNewWorkbook)
{
	if (!vData.IsNull()) throw ExcelAppAttachedException("CreateApp");

	if (nSheetsInNewWorkbook < 1) nSheetsInNewWorkbook = 1;
	if (vData.IsNull()) {
		vData = Variant::CreateObject("Excel.Application");
		vData.OlePropertySet("Visible", visible);
		vData.OlePropertySet("SheetsInNewWorkbook", nSheetsInNewWorkbook);
	}
	return this;
}
/// @}

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

	/// Удалить всю дочернюю структуру, т.к. больше нет связи, дочерние элементы
	/// не актуальны.
	ClearChilds();
}

/// @details Освобождает экземпляр эксель и вычищает себя из памяти
void TExcelApp::Free(){
	vData.OlePropertySet("DisplayAlerts", false);
	vData.OleProcedure("Quit");
	vData.OlePropertySet("Interactive", true);
	//delete this;
	this->~TExcelApp();
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

/// @details Открывает книгу по запрашиваемому пути
/// @param path Путь до файла
TExcelWorkbook* TExcelApp::OpenWorkbook(const String& path)
{
	if (vData.IsNull()) throw ExcelAppNotAttachedException("CreateWorkbook");

	vData.OlePropertyGet("Workbooks").OleProcedure("Open", System::StringToOleStr(path));
	seekAndSetDataChild("Workbooks", getChildCountByType("Workbooks"));

	TExcelWorkbook* out = new TExcelWorkbook(this, vDataChild);
	return out;
}

}


