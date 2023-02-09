#ifndef uTExcelAppH
#define uTExcelAppH

#include "uTExcelWorkbook.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс для работы с приложением Excel
 * @ingroup ExcelClientObjects
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelApp :
	public TExcelObjectTemplate<TExcelApp>
{
public:
	TExcelApp();
    TExcelApp(const TExcelApp& src);
    ~TExcelApp();

	TExcelApp(bool visible);  
	TExcelApp(bool visible, unsigned int nSheetsInNewWorkbook);

private:
	bool Notifications; ///< Включены ли оповещения

	void Init(); ///< Метод первичной иницилизации

public:
	TExcelApp* CreateApp(bool visible);
	TExcelApp* CreateApp(bool visible, unsigned int nSheetsInNewWorkbook);
	
	TExcelApp* AttachApp(); ///< Подключиться к приложению
	void DeattachApp(); ///< Отключиться
	TExcelApp* TryAttachApp(); ///< Попытаться подключиться, или создать экземпляр нового

	void Close(bool silent = true); ///< Закрыть штатно
	void Free(); ///< Освободить кабанчиком

	TExcelApp* SetExcelNotifications(bool stat); ///< Установить режим отображения сообщений
	TExcelApp* SetSheetsInNewWorkbook(unsigned int N); ///< Установить количество новых страниц

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

