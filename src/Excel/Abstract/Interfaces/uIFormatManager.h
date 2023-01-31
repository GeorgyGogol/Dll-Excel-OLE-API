#ifndef UIFORMATMANAGER_H
#define UIFORMATMANAGER_H

// Интерфейс для настройки формата ячеек

template<class T>
class IFormatManager {
protected:
	IFormatManager<T>() {}
	virtual ~IFormatManager() = 0;

public:
	virtual T* SetHorizontalAlign(ExcelTextAlign align) = 0;
	virtual T* SetVerticalAlign(ExcelTextAlign align) = 0;
	 
	virtual T* SetBorders() = 0;
	 
	virtual T* SetWidth() = 0;
	virtual T* SetHeight() = 0;
	virtual T* AutoSize() = 0;
	 
	virtual T* SetFormat() = 0;	
};

#endif

