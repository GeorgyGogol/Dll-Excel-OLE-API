//---------------------------------------------------------------------------

#include "EntryPointDLL.h"


//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void* lpReserved)
{
	CoInitialize(NULL);
	return 1;
}
//---------------------------------------------------------------------------