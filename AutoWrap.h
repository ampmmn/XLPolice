#pragma once

#include <windows.h>
#include <stdint.h>

HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...);

class DispWrapper
{
public:
	DispWrapper();
	DispWrapper(IDispatch* disp);
	~DispWrapper();

	void Release();

	IDispatch** operator &();
	operator IDispatch*();

	int64_t GetPropertyInt64(LPCOLESTR name);
	bool GetPropertyObject(LPCOLESTR name, DispWrapper& object);

	void CallVoidMethod(LPCOLESTR methodName);

private:
	IDispatch* mDispPtr;
};
