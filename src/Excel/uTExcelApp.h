//---------------------------------------------------------------------------

#ifndef uTExcelAppH
#define uTExcelAppH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "uTExcelWorkbook.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelApp : public TExcelObject
{
public:
	TExcelApp();
    // И сразу иницилизировать приложение
	TExcelApp(bool visible);  
	TExcelApp(bool visible, unsigned int nSheetsInNewWorkbook);
    
    TExcelApp(const TExcelApp& src);
    ~TExcelApp();

private:
	bool Notifications;

	void Init();

public:
    // Создать/подключиться к приложению
	TExcelApp* CreateApp(bool visible);
	TExcelApp* CreateApp(bool visible, unsigned int nSheetsInNewWorkbook);
	TExcelApp* TryAttachApp();
	TExcelApp* AttachApp();
    // Отключится
	void DeattachApp();
    // Закрыть
	void Close(bool silent = true);
	//void Free(); // ???

	TExcelApp* SetExcelNotifications(bool stat);
	TExcelApp* SetSheetsInNewWorkbook(unsigned int N);

	// Workbook - методы-обертки для объектов ниже по иерархии
    unsigned int WorkbookCount();
	TExcelWorkbook* CreateWorkbook();
	TExcelWorkbook* CreateWorkbook(const String& workbookName);
	TExcelWorkbook* GetCurrentWorkbook();
    //TExcelWorkbook* ActiWorkbook(const String& workbookName);
    //TExcelWorkbook* ActiWorkbook(unsigned int N);
};

//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif


