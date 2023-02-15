//---------------------------------------------------------------------------

#ifndef uCheckersH
#define uCheckersH

#include "Classes.hpp"
//---------------------------------------------------------------------------
namespace checkers {

/// Проверка одного значения
/// @{
template<typename Tv>
bool checkNumber(Tv& n)
{
	if (n < 1) n = 1;
	return true;
}
//bool checkNumber(int& n);
/// @}

/// Проверка пары значений
/// @{
//bool checkPairNumbers(unsigned int& c, unsigned int& r);
//bool checkPairNumbers(int& c, int& r);
/// @}

/// Проверка Варанты на содержание объекта
//bool checkVariantOnNull(const Variant& var);

}
//---------------------------------------------------------------------------
#endif
