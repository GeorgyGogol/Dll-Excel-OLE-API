//---------------------------------------------------------------------------

#ifndef uTLogH
#define uTLogH

//---------------------------------------------------------------------------

#include "uLogCallReasons.h"

class TLog 
{
public:
	TLog();
	TLog(CallLogReason reason);
	TLog(const char* objectName);
	~TLog();

private:
	char* Prefix;

	void logInit();
	char* getTextTemplate(CallLogReason templ) const;

public:
	void WriteFileStamp();
	void WriteSize(const unsigned int& size);

	void WriteLog(const char* message);
	template<typename T> void WriteLog(const char* message, const T& value);
	
	void WriteLog(const char* message, CallLogReason reason);

};

//---------------------------------------------------------------------------
#endif
