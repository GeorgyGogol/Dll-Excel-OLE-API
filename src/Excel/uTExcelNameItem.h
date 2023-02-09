#ifndef uTExcelNameItemH
#define uTExcelNameItemH

#include "uTExcelObjectTemplate.h"
//---------------------------------------------------------------------------
namespace exl {
/** -------------------------------------------------------------------------
 * @brief Класс именованных диапазонов
 * @ingroup ExcelClientObjects
 * 
 * Служит для взаимодействия с именованным диапазоном.
 * 
 * ---------------------------------------------------------------------- **/
class DLL_EI TExcelNameItem : public TExcelObjectTemplate<TExcelNameItem>
{
private:
    //TExcelNameItem();

public:
    TExcelNameItem(TExcelObject* pParent, const Variant& data);
    TExcelNameItem(const TExcelNameItem& src);
    virtual ~TExcelNameItem();

public:
	TExcelNameItem* SetValue(const Variant& val); ///< Записать значение
	TExcelNameItem* SetValue(const String& val); ///< Записать строку
	TExcelNameItem* SetRefers(const String& address); ///< Записать формулу или ссылку

};
/** -------------------------------------------------------------------------
 * @brief Интерфейс управления именованными диапазонами
 * @ingroup Interfaces
 * 
 * Предназначен для контроля реализацией методов взаимодействия с именованными
 * диапазонами.
 * 
 * ---------------------------------------------------------------------- **/
class ITExcelNames {
public:
	/// @brief Найти и получить именованный объект
	/// @param itemName Название объекта
	/// @return Именованный объект
	/// @note Объекты упорядочены по алфавиту
	virtual TExcelNameItem* GetNameItem(const String& itemName) = 0;
	/// @param N Номер
	virtual TExcelNameItem* GetNameItem(unsigned int N) = 0;

	/// @brief Создать именованный объект
	/// @param itemName Название
	/// @return Созданный именованный объект
	virtual TExcelNameItem* AddNamedItem(const String& itemName) = 0;

};

}

//---------------------------------------------------------------------------
#endif
