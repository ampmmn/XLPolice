#include "ExcelProxy.h"
#include "AutoWrap.h"

struct ExcelProxy::PImpl
{
	DispWrapper mApp;
};

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

static bool GetExcelApplication(DispWrapper& excelApp)
{
		CLSID clsid;
		HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
		if (FAILED(hr)) {
			return false;
		}

		IUnknown* unkPtr = nullptr;
		hr = GetActiveObject(clsid, NULL, &unkPtr);
		if(FAILED(hr)) {
			// Excel is not running.
			return false;
		}

		hr = unkPtr->QueryInterface(&excelApp);
		unkPtr->Release();

		return SUCCEEDED(hr);
}

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

ExcelProxy::ExcelProxy() : in(new PImpl)
{
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		return;
	}
	GetExcelApplication(in->mApp);
}

ExcelProxy::~ExcelProxy()
{
	if (((IDispatch*)in->mApp) != nullptr) {
		in->mApp.Release();
	}

	CoUninitialize();
}


bool ExcelProxy::IsAvailable()
{
	return ((IDispatch*)in->mApp) != nullptr;
}

bool ExcelProxy::HasDocument()
{
	DispWrapper activeSheet;
	if (in->mApp.GetPropertyObject(L"ActiveSheet", activeSheet) == false) {
		return false;
	}

	if (((IDispatch*)activeSheet) != nullptr) {
		return true;
	}

	HWND hwndApp = (HWND)in->mApp.GetPropertyInt64(L"Hwnd");
	if (IsWindow(hwndApp) == false) {
		return false;
	}

	LONG_PTR style = GetWindowLongPtr(hwndApp, GWL_STYLE);
	return (style & WS_VISIBLE) != 0;
}

void ExcelProxy::Terminate()
{
	DWORD pid = 0;
	HWND hwndApp = (HWND)in->mApp.GetPropertyInt64(L"Hwnd");
	if (IsWindow(hwndApp)) {
		GetWindowThreadProcessId(hwndApp, &pid);
	}

	in->mApp.CallVoidMethod(L"Quit");

	if (pid == 0) {
		return;
	}

	HANDLE h = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_TERMINATE, FALSE, pid);

	ULONGLONG start = GetTickCount64();
	while(GetTickCount64() - start <= 2000) {
		DWORD exitCode = 0;
		if (GetExitCodeProcess(h, &exitCode) && exitCode != STILL_ACTIVE) {
			return;
		}
		Sleep(100);
	}
	TerminateProcess(h, 0);
}
