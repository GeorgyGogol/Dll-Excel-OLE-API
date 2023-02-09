#ifndef uTExcelExceptionsH
#define uTExcelExceptionsH
/** @file 
 * @ingroup Exceptions
 * @brief Базовое исключение
 * @details Самый первый предок, породивший всех остальных.
 */

/**
 * @defgroup Exceptions Исключения
 * @brief Ошибки-ачивки, выпадающие во время работы
 * @details Структура объектов-наследников класса Exception, выпадающие когда что-то
 * пошло не так.
 * @{
 * @}
 */

#include "uDll.h"
//---------------------------------------------------------------------------
/** -------------------------------------------------------------------------
 * @brief Базовый класс исключений
 * @ingroup Exceptions
 * 
 * Содержит все извращения, которые нужны для генерации исключительной 
 * ситуации.
 * 
 * ---------------------------------------------------------------------- **/
class ExcelException : public Exception
{
public:
	ExcelException(const String& message);
	ExcelException(const char* message);

protected:
	ExcelException(const String& classWhere, const String& method, const String& message);
	ExcelException(const char* classWhere, const char* method, const char* message);

private:
	void WriteError(const char* error);

};

//---------------------------------------------------------------------------
#endif

