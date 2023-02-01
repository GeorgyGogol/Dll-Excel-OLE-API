/** @file 
 * @brief Самый главный заголовок для всех вайлов
 * 
 * Здесь включаются зависимости, настройки, определяется макрос для экспортируемых
 * объектов. Также здесь идет включение перечислений, ну мало ли... Изначально
 * перечисления были вообще одним файлом.
 */
#ifndef uDllH
#define uDllH

//---------------------------------------------------------------------------
#include "uDependeces.h"
#include "uBuildSettings.h"

#ifdef __DLL__
#  define DLL_EI __declspec(dllexport)
#else
#  define DLL_EI __declspec(dllimport)
#endif

#include "uTExcelEnums.h"

//---------------------------------------------------------------------------
#endif
