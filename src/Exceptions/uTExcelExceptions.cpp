//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelExceptions.h"
#include "Log.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
/// @details Открытый конструкторы, для возможности сброса исключения с только
/// сообщением.
/// @param message Сообщение об ошибке
ExcelException::ExcelException(const String& message)
	: Exception(message)
{
	WriteError(Message.c_str());
}

/// @overload
ExcelException::ExcelException(const char* message)
    : Exception(message)
{
	WriteError(Message.t_str());
}

/// @details Конструкторы, которые используются для генерации исключений с
/// более подробным описанием (дополнительно: класс, который сбрасывает и
/// метод этого класса, где произошло).
/// @param classWhere Класс, где произошло
/// @param method Какой метод
/// @param message Сообщение
ExcelException::ExcelException(const String& classWhere, const String& method, const String& message)
    : Exception(classWhere + "::" + method + ": " + message)
{
	WriteError(Message.t_str());
}

/// @overload
ExcelException::ExcelException(const char* classWhere, const char* method, const char* message)
    : Exception(String(classWhere) + "::" + String(method) + ": " + String(message))
{
	WriteError(Message.t_str());
}

/// @details Делает соответствующую запись в лог
/// @param error Сообщение об ошибке
void ExcelException::WriteError(const String& error)
{
    TLog* log = new TLog(CallLogReason::Error);
    log->WriteLog(error.t_str());
    delete log;
}

