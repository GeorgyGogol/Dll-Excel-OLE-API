//---------------------------------------------------------------------------

#ifndef uTExcelAppH
#define uTExcelAppH

//---------------------------------------------------------------------------
#include "uTExcelWorkbook.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
class DLL_EI TExcelApp : public TExcelObjectTemplate<TExcelApp>
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
	//TExcelApp* SetWorkfolder(const String& path);

    // Создать экземпляр приложения
	TExcelApp* CreateApp(bool visible);
	TExcelApp* CreateApp(bool visible, unsigned int nSheetsInNewWorkbook);
	
	// Подключится/отключиться
	TExcelApp* AttachApp();
	void DeattachApp();
	
	// Попробовать подключится, иначе - создать
	TExcelApp* TryAttachApp();

	// Закрыть 
	void Close(bool silent = true); // Штатно
	void Free(); // Внештатно

	// Настройки приложения
	TExcelApp* SetExcelNotifications(bool stat);
	TExcelApp* SetSheetsInNewWorkbook(unsigned int N);

	// Всё, что Workbook касается - породить или получить текущую книгу
    unsigned int WorkbookCount();
	TExcelWorkbook* CreateWorkbook();
	TExcelWorkbook* CreateWorkbook(const String& workbookName);
	TExcelWorkbook* GetCurrentWorkbook();

	TExcelWorkbook* OpenWorkbook(const String& path);
};

}
//---------------------------------------------------------------------------
#endif

