//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTLog.h"
#include <fstream>
#include "Classes.hpp"

#include "uBuildSettings.h"

#ifndef LOG_FILE_NAME
#define LOG_FILE_NAME "ExcelUseLog.log"
#endif

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
using namespace std;

TLog::TLog() :
	Prefix(NULL)
{
	logInit();
}

TLog::TLog(CallLogReason reason) :
	Prefix(NULL)
{
	logInit();

	switch (reason)
	{
	case CallLogReason::Error:
		WriteLog("Error thrown!");
		break;

	case CallLogReason::Warning:
		WriteLog("Warning!");
		break;
	
	case CallLogReason::None:
	default:
		break;
	}
}

TLog::TLog(const char* objectName)
{
	logInit();

	Prefix = const_cast<char*>(objectName);
}

TLog::~TLog()
{}

void TLog::logInit()
{
	if (!FileExists(LOG_FILE_NAME)) {
		ofstream log;
		log.open(LOG_FILE_NAME, ios::trunc);
		log.close();
		WriteFileStamp();
	}
}

char* TLog::getTextTemplate(CallLogReason reason) const
{
	char* out = NULL;

	switch (reason)
	{
	case CallLogReason::Error:
		out = "Error thrown: %s";
		break;
	
	case CallLogReason::Warning:
		out = "Warning: %s";
		break;
	
	case CallLogReason::WriteUsedMemory:
		out = "Memory used: %.2f %s";
		break;

	case CallLogReason::None:
	default:
		break;
	}

	return out;
}

void TLog::WriteFileStamp()
{
	ofstream log;
	String dtStamp = Now();

	log.open(LOG_FILE_NAME, ios::app);
	if (log.is_open()) {
		for (auto i = 0; i < 72; i++) log<<"-";
		log<<"\nLog start: "<<dtStamp.t_str()<<endl;
		log<<endl;
		log.close();
	}
}

void TLog::WriteSize(const unsigned int& size)
{
	char* cTemplate = getTextTemplate(CallLogReason::WriteUsedMemory);
	double dSize = static_cast<double>(size);
	char* cMulty = "Bytes";

	if (dSize > 1024.0 * 5) {
		dSize /= 1024;
		cMulty = "KBytes";
	}
	if (dSize > 1024.0 * 5) {
		dSize /= 1024;
		cMulty = "MBytes";
	}

	AnsiString logMessage;
	logMessage.printf(cTemplate, dSize, cMulty);

	WriteLog(logMessage.c_str());
}

void TLog::WriteLog(const char* message)
{
	ofstream log;
	String dtStamp = Now();

	log.open(LOG_FILE_NAME, ios::app);
	if (log.is_open()) {

		log<<dtStamp.t_str()<<": ";
		if (Prefix) log<<Prefix<<": ";
		log<<message;
		log<<endl;

		log.close();
	}
}


template<typename T>
void TLog::WriteLog(const char* message, const T& value)
{
	ofstream log;
	String dtStamp = Now();

	log.open(LOG_FILE_NAME, ios::app);
	if (log.is_open()) {

		log<<dtStamp.t_str()<<": ";
		if (Prefix) log<<Prefix<<": ";
		log<<message<<": ";
		log<<value;
		log<<endl;

		log.close();
	}
}

void TLog::WriteLog(const char* message, CallLogReason reason)
{
	if (reason == CallLogReason::None) return;

	char* cTemplate = getTextTemplate(reason);
	AnsiString logMessage;
	logMessage.printf(cTemplate, message);

	WriteLog(logMessage.c_str());
}


template void TLog::WriteLog<>(const char*, const unsigned int&);
template void TLog::WriteLog<>(const char*, const int&);
template void TLog::WriteLog<>(const char*, const double&);

