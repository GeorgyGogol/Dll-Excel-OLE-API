﻿//---------------------------------------------------------------------------

#ifndef uTExcelEnumsH
#define uTExcelEnumsH

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
#include "Enums/uTExcelEnumsPivot.h"

//---------------------------------------------------------------------------
// Выравниывние ячеек (не все возможные, есть еще куча всякого...)
enum class ExcelTextAlign : short int {
	xlCenter = -4108,	// Центр
	xlBottom = -4107,	// Низ
	xlTop = -4160,		// Верх
	xlLeft = -4131,		// Лево
	xlRight = -4152		// Право
};

// Направление для заполнения (?)
enum class FillDirection : char {
	Down = 0, Up, Left, Right
};

}
//---------------------------------------------------------------------------
#endif  // uTExcelEnumsH
