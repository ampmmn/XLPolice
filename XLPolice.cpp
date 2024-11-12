#include "framework.h"
#include "ExcelProxy.h"

int APIENTRY wWinMain(
	_In_ HINSTANCE,
	_In_opt_ HINSTANCE ,
	_In_ LPWSTR,
	_In_ int
)
{
	int count = 0;
	for(;;) {
		ExcelProxy excel;
		if (excel.IsAvailable() && excel.HasDocument() == false) {
			count++;
		}
		else {
			count = 0;
		}

		if (count >= 10) {
			excel.Terminate();
			count = 0;
		}

		Sleep(1000);
	}
	return 0;
}

