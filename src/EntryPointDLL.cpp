//---------------------------------------------------------------------------

#include <vcl.h>
#include <windows.h>
#pragma hdrstop

//---------------------------------------------------------------------------
#pragma argsused
#include "ExcelAPI.h"
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void* lpReserved)
{
	CoInitialize(NULL);
	return 1;
}
//---------------------------------------------------------------------------


