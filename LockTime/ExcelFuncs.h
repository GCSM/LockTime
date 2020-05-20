// XLFunctions.h
#ifndef XLFunctions_h
#define XLFunctions_h
#include <windows.h>

HRESULT SetVisible(IDispatch* pObject, LCID lcid);
HRESULT GetXLCell(IDispatch* pXLWorksheet, LCID& lcid, wchar_t* pszRange, wchar_t* pszCell, size_t iBufferLength);
HRESULT GetCell(IDispatch* pXLSheet, LCID lcid, wchar_t* pszRange, VARIANT& pVt);
IDispatch* SelectWorkSheet(IDispatch* pXLWorksheets, LCID& lcid, wchar_t* pszSheet);
IDispatch* OpenXLWorkBook(IDispatch* pXLWorkbooks, LCID& lcid, wchar_t* pszWorkBookPath);
IDispatch* GetDispatchObject(IDispatch* pCallerObject, DISPID dispid, WORD wFlags, LCID lcid);

#endif 