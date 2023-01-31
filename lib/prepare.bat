:: Файл подготовки библиотек, чтобы все и сразу

@echo off

:: Перемещение lib
copy /Y bpi-lib_out\Excel.lib

:: Сжатие dll
if exist upx_32 (
upx_32 -9 Release-Win32\ExcelAPI.dll
upx_32 -9t Release-Win32\ExcelAPI.dll
) 
if exist upx (
upx -9 Release-Win32\ExcelAPI.dll
upx -9t Release-Win32\ExcelAPI.dll
) else (
echo UPX not found! Compress is not done!
)

Pause