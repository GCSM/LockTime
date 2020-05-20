// XLFunctions.cpp
#ifndef   UNICODE
#define   UNICODE
#endif
#ifndef   _UNICODE
#define   _UNICODE
#endif
#include <stdio.h>
#include "ExcelFuncs.h"

HRESULT SetVisible(IDispatch* pObject, LCID lcid)
{
	VARIANT vArgArray[1];
	DISPPARAMS DispParams;
	DISPID dispidNamed;
	VARIANT vResult;
	HRESULT hr;

	VariantInit(&vArgArray[0]);
	vArgArray[0].vt = VT_BOOL;
	vArgArray[0].boolVal = TRUE;
	dispidNamed = DISPID_PROPERTYPUT;
	DispParams.rgdispidNamedArgs = &dispidNamed;    // fact of the matter is that the direct VTable part of the dual interface is largely nonfunctional in all the MS Office applications
	DispParams.cArgs = 1;                           // of which I'm familiar, specifically Word and Excel.  So this horrendous IDispatch access is forced upon us.  Its only redeeming
	DispParams.cNamedArgs = 1;                      // feature is that it does work in spite of the fact that it is horribly awkward.
	VariantInit(&vResult);
	hr = pObject->Invoke(0x0000022e, IID_NULL, lcid, DISPATCH_PROPERTYPUT, &DispParams, &vResult, NULL, NULL);
	return hr;
}