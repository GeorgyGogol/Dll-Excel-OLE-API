#ifndef uTExcelTableCreatorExceptionH
#define uTExcelTableCreatorExceptionH
/** @file 
 * @brief Исключение помошника создания таблиц
 * @ingroup Exceptions
 */

#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
/// @addtogroup Exceptions
/// @{
/**
 * @brief Базовый исключений помошника создания таблиц
 * 
 * */
class ExcelTableCreatorException : public ExcelException
{
public:
	ExcelTableCreatorException(const String& method, const String& message);
};
/// @}

//---------------------------------------------------------------------------
#endif

