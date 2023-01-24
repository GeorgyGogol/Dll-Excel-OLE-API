@echo off
copy /Y bpi-lib_out\Excel.lib
upx_32 -9 Release-Win32\ExcelAPI.dll
upx_32 -9t Release-Win32\ExcelAPI.dll
Pause