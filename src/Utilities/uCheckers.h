//---------------------------------------------------------------------------

#ifndef uCheckersH
#define uCheckersH

#include "Classes.hpp"
//---------------------------------------------------------------------------
namespace checkers {

/// �������� ������ ��������
/// @{
template<typename Tv>
bool checkNumber(Tv& n)
{
	if (n < 1) n = 1;
	return true;
}
//bool checkNumber(int& n);
/// @}

/// �������� ���� ��������
/// @{
//bool checkPairNumbers(unsigned int& c, unsigned int& r);
//bool checkPairNumbers(int& c, int& r);
/// @}

/// �������� ������� �� ���������� �������
//bool checkVariantOnNull(const Variant& var);

}
//---------------------------------------------------------------------------
#endif
