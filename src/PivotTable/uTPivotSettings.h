//---------------------------------------------------------------------------

#ifndef uTPivotSettingsH
#define uTPivotSettingsH

//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------
#include "../Table/uTTableCreator.h"
//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
// Размещение (агрегация) данных в сводной таблице
enum PivotDataPlace : char { Rows = 0, Columns = 1 };

//---------------------------------------------------------------------------
// Настройка - столбик и его агрегация
struct DLL_EI TExcelTablePivotField
{
	String ColumnName;
	String Caption;

	XlPivotFieldOrientation Type;
	XlConsolidationFunction Function;
};

//---------------------------------------------------------------------------
// Класс-контейнер для настройки сводной таблицы
class DLL_EI TPivotSettings
{
public:
	TPivotSettings(const String& pivotName);
	~TPivotSettings();

	typedef std::vector<TExcelTablePivotField> tSettingsArray;

private:
	String PivotName;		// Название таблицы

public:
	String NewSheetName;
	
	String GetPivotName();
	void SetPivotName(const String& name);

	String PivotTitle;		// Заголовок
	tSettingsArray Settings;	// Массив настроек, да в общем доступе
	PivotDataPlace DataPlace; // Расположение значений - по строчкам или столбикам
	bool ShowRowTotal;		// Показать Сумму по столбику
	bool ShowColumnTotal;	// Показать Сумму по столбцу

	void Clear(); // Очистить настройки
	void Add(const TExcelTablePivotField& setting);	// Добавить 
	void AddRow(const String& сolumnName, const String& caption = "");
	void AddColumn(const String& сolumnName, const String& caption = "");
	void AddData(const String& сolumnName, const String& caption = "",
		XlConsolidationFunction function = XlConsolidationFunction::xlUnknown);

	unsigned int RowCount() const;
	unsigned int ColumnCount() const;
};
//---------------------------------------------------------------------------
}
//---------------------------------------------------------------------------
#endif

