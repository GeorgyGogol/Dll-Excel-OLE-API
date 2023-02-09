#ifndef uTExcelAppExceptionsH
#define uTExcelAppExceptionsH
/** @file 
 * @brief Исключения приложения-хоста
 * @ingroup Exceptions
 */

#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
/**
 * @addtogroup Exceptions
 * @{
 * 
 * @brief Базовый класс для исключений приложения
 * 
 * Привязан к TExcelApp
 *
*/
class ExcelAppException : public ExcelException
{
public:
	ExcelAppException(const char* method, const char* message);
};

//---------------------------------------------------------------------------
/**
 * @brief Ошибка взаимодействия с неиницилизированным приложением
 * 
 */
class ExcelAppNotAttachedException : public ExcelAppException
{
public:
	ExcelAppNotAttachedException(const char* method);
};

//---------------------------------------------------------------------------
/**
 * @brief Ошибка, которая возникает при попытке подключиться еще раз
 * 
 */
class ExcelAppAttachedException : public ExcelAppException
{
public:
	ExcelAppAttachedException(const char* method);
};
/// @}

//---------------------------------------------------------------------------
#endif

