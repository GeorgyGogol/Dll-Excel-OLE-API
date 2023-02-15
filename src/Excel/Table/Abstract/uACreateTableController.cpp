//---------------------------------------------------------------------------


#pragma hdrstop

#include "uACreateTableController.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

//---------------------------------------------------------------------------
namespace exl {
//---------------------------------------------------------------------------
ACreateTableController::ACreateTableController() : 
	NeedDisableDataSet(false)
{}

void ACreateTableController::setNeedDisableDataSet(bool isNeedDisable)
{
    NeedDisableDataSet = isNeedDisable;
}

}

