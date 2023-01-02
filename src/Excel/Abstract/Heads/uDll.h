//---------------------------------------------------------------------------

#ifndef uDllH
#define uDllH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------

#ifdef __DLL__
#  define DLL_EI __declspec(dllexport)
#else
#  define DLL_EI __declspec(dllimport)
#endif

#include <Classes.hpp>

#include "uTExcelEnums.h"

//---------------------------------------------------------------------------
#endif

