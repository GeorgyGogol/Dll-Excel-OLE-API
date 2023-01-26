//---------------------------------------------------------------------------

#include "EntryPointDLL.h"

//---------------------------------------------------------------------------
#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void* lpReserved)
{
	CoInitialize(NULL);
	return 1;
}
//---------------------------------------------------------------------------


