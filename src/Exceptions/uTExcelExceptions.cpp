//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelExceptions.h"
#include "Log.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
/// @{
/// @param message Сообщение об ошибке
ExcelException::ExcelException(const String& message)
	: Exception(message)
{
	WriteError(Message.t_str());
}

ExcelException::ExcelException(const char* message)
    : Exception(message)
{
	WriteError(Message.t_str());
}
/// @}

/// @{
/// @param classWhere Класс, где произошло
/// @param method Какой метод
/// @param message Сообщение
ExcelException::ExcelException(const String& classWhere, const String& method, const String& message)
    : Exception(classWhere + "::" + method + ": " + message)
{
	WriteError(Message.t_str());
}

ExcelException::ExcelException(const char* classWhere, const char* method, const char* message)
    : Exception(String(classWhere) + "::" + String(method) + ": " + String(message))
{
	WriteError(Message.t_str());
}
/// @}

void ExcelException::WriteError(const char* error)
{
    TLog* log = new TLog(CallLogReason::Error);
    log->WriteLog(error);
    delete log;
}

