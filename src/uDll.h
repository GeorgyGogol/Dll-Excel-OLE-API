#ifndef uDllH
#define uDllH
/** @file 
 * @brief Самый главный заголовок для всех файлов
 * 
 * Здесь включаются зависимости, настройки, определяется макрос для экспортируемых
 * объектов. Также здесь идет включение перечислений, ну мало ли... Изначально
 * перечисления были вообще одним файлом.
 */

//---------------------------------------------------------------------------
#include "uDependeces.h"
#include "uBuildSettings.h"

#ifdef __DLL__
#  define DLL_EI __declspec(dllexport)
#else
#  define DLL_EI __declspec(dllimport)
#endif

#include "uEnums.h"

//---------------------------------------------------------------------------
#endif
