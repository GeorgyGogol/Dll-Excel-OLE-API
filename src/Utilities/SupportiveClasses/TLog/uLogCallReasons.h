#ifndef ULOGCALLREASONS_H
#define ULOGCALLREASONS_H

//---------------------------------------------------------------------------

enum CallLogReason : unsigned short int {
	None = 0,
	Error,
	Warning,
	WriteUsedMemory,
};

//---------------------------------------------------------------------------
#endif // ULOGCALLREASONS_H

