#include "AutoWrap.h"
#include <vector>

// 	https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c
HRESULT AutoWrap(
	int autoType,
	VARIANT* pvResult,
	IDispatch* pDisp,
	LPCOLESTR ptName,
	int cArgs...
)
{
	if (!pDisp) {
		return E_FAIL;
	}

	va_list marker;
	va_start(marker, cArgs);

	DISPID dispID;
	HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		return hr;
	}

	std::vector<VARIANT> args(cArgs + 1);
	for (int i = 0; i < cArgs; i++) {
		args[i] = va_arg(marker, VARIANT);
	}
	va_end(marker);

	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	dp.cArgs = cArgs;
	dp.rgvarg = &args.front();

	DISPID dispidNamed = DISPID_PROPERTYPUT;
	if (autoType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	return pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, (WORD)autoType, &dp, pvResult, NULL, NULL);
}

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

DispWrapper::DispWrapper(): mDispPtr(nullptr)
{
}

DispWrapper::DispWrapper(IDispatch* disp) : mDispPtr(disp)
{
	if (disp) {
		disp->AddRef();
	}
}

DispWrapper::~DispWrapper()
{
	Release();
}

void DispWrapper::Release()
{
	if (mDispPtr) {
		mDispPtr->Release();
		mDispPtr = nullptr;
	}
}

IDispatch** DispWrapper::operator &()
{
	return &mDispPtr;
}

DispWrapper::operator IDispatch*()
{
	return mDispPtr;
}


int64_t DispWrapper::GetPropertyInt64(
		LPCOLESTR name
)
{
	VARIANT result;
	VariantInit(&result);

	AutoWrap(DISPATCH_PROPERTYGET, &result, mDispPtr, name, 0);
	return result.llVal;
}

bool DispWrapper::GetPropertyObject(LPCOLESTR name, DispWrapper& object)
{
	VARIANT result;
	VariantInit(&result);

	HRESULT hr = AutoWrap(DISPATCH_PROPERTYGET, &result, mDispPtr, name, 0);
	if (FAILED(hr)) {
		return false;
	}
	object = result.pdispVal;

	return true;
}

void DispWrapper::CallVoidMethod(LPCOLESTR methodName)
{
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_METHOD, &result, mDispPtr, methodName, 0);
}


