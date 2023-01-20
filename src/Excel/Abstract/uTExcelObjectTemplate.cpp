//---------------------------------------------------------------------------


#pragma hdrstop

#include "uTExcelObjectTemplate.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
template<class T>
TExcelObjectTemplate<T>::TExcelObjectTemplate()
    : TExcelObject()
{
}

template<class T>
TExcelObjectTemplate<T>::TExcelObjectTemplate(TExcelObject* pParent, const Variant& data)
    : TExcelObject(pParent, data)
{
}

template<class T>
TExcelObjectTemplate<T>::TExcelObjectTemplate(const TExcelObject& src)
    : TExcelObject(src)
{
}

template<class T>
TExcelObjectTemplate<T>::~TExcelObjectTemplate()
{
}

template<class T>
T* TExcelObjectTemplate<T>::Show() {
    return (T*)TExcelObject::Show();
}

template<class T>
T* TExcelObjectTemplate<T>::Hide() {
    return (T*)TExcelObject::Hide();
}

template<class T>
T* TExcelObjectTemplate<T>::SetName(const String& newName) {
    return (T*)TExcelObject::SetName(newName);
}

// Для каждого типа объедка - нужон шаблон
class DLL_EI TExcelApp;
template class TExcelObjectTemplate<TExcelApp>;

class DLL_EI TExcelWorkbook;
template class TExcelObjectTemplate<TExcelWorkbook>;

class DLL_EI TExcelSheet;
template class TExcelObjectTemplate<TExcelSheet>;

class DLL_EI TExcelCells;
template class TExcelObjectTemplate<TExcelCells>;

class DLL_EI TExcelTable;
template class TExcelObjectTemplate<TExcelTable>;

}

