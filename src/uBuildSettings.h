#ifndef BUILDSETTINGS_H
#define BUILDSETTINGS_H
/** @file 
 * @brief Файл настроек сборки библиотеки
 * 
 * Ну, тут вроде бы все понятно.
 * 
 * Настройки можно посмотреть в doc/BuildSettings.md
 */

// Общие настройки


/// Переобределить имя/путь файла с логами
//#define LOG_FILE_NAME "ExcelUseLog.log"

// Настройки для Debug
#ifdef _DEBUG

#define ENABLE_USAGE_STATISTIC

#endif // _DEBUG


#endif // BUILDSETTINGS_H
