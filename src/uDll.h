// ---------------------------------------------------------------------------

#ifndef uDllH
#define uDllH

// ---------------------------------------------------------------------------

#include "uDependeces.h"

#ifdef __DLL__
#  define DLL_EI __declspec(dllexport)
#else
#  define DLL_EI __declspec(dllimport)
#endif

#include "uTExcelEnums.h"

// ---------------------------------------------------------------------------
#endif
