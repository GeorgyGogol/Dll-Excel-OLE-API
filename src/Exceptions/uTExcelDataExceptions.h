#ifndef uTExcelDataExceptionsH
#define uTExcelDataExceptionsH
/** @file 
 * @brief Исключение варианты
 * @ingroup Exceptions
 */

#include "uTExcelExceptions.h"
//---------------------------------------------------------------------------
/**
 * @addtogroup Exceptions
 * @{
 * 
 * @brief Базовый класс исключения варианты
 * 
 * Сбрасывается, если что-то не то происходит с вариантой, где живет объект 
 * Excel. Например, объекта там нет и отправлять команду - некому.
 * 
*/
class ExcelDataException : public ExcelException
{
public:
	ExcelDataException(const String& message);
	ExcelDataException(const char* method, const char* message);
	ExcelDataException(const char* message);
};

//---------------------------------------------------------------------------
/**
 * @brief Варианта не иницилизирована
 * 
 */
class ExcelDataNullException : public ExcelDataException
{
public:
	ExcelDataNullException();
};

//---------------------------------------------------------------------------
/**
 * @brief 
 * 
 */
/* class ExcelSelectCellException : public ExcelDataException
{
public:
	ExcelSelectCellException();
	ExcelSelectCellException(const char* type);
}; */
/// @}

//---------------------------------------------------------------------------
#endif

