#include <windows.h>
#include <sddl.h>
#include <stdio.h>
#include <winevt.h>
#include <cstdint>
#include "ExcelFuncs.h"


#pragma comment(lib, "wevtapi.lib")
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "kernel32.lib")

#import "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16\mso.dll" rename( "RBG", "MSORGB")  //Order seems to matter with these directives

using namespace Office;

#import "C:\\Program Files\\Microsoft Office\\root\vfs\\ProgramFilesCommonX86\\Microsoft Shared\VBA\VBA6\VBE6EXT.olb"


using namespace VBIDE;

#import "C:\Program Files\Microsoft Office\root\Office16\Excel.exe" \
    rename( "DialogBox", "ExcelDialogBox" ) \
    rename( "RGB", "ExcelRGB" ) \
    rename( "CopyFile", "ExcelCopyFile" ) \
    rename( "ReplaceText", "ExcelReplaceText" ) \
    exclude( "IFont", "IPicture" ) no_dual_interfaces


#define _SECOND ((__int64) 10000000)
#define _MINUTE (60 * _SECOND)
#define _HOUR   (60 * _MINUTE)
#define _DAY    (24 * _HOUR)


#define ARRAY_SIZE 10
#define TIMEOUT 1000  // 1 second; Set and use in place of INFINITE in EvtNext call

DWORD GetResults(EVT_HANDLE hResults, SYSTEMTIME& ststart, SYSTEMTIME& stend);

FILETIME calcStart() {

	ULONGLONG qwResult;
	SYSTEMTIME stCurrentTime;
	FILETIME ftCurrentTime;

	GetSystemTime(&stCurrentTime);
	SystemTimeToFileTime(&stCurrentTime, &ftCurrentTime);
	qwResult = (((ULONGLONG)ftCurrentTime.dwHighDateTime) << 32) + ftCurrentTime.dwLowDateTime;  //Put current system time into quad word

	qwResult -= 1 * _DAY;  //Yesterday

	ftCurrentTime.dwLowDateTime = (DWORD)(qwResult & 0xFFFFFFFF);
	ftCurrentTime.dwHighDateTime = (DWORD)(qwResult >> 32);

	FileTimeToSystemTime(&ftCurrentTime, &stCurrentTime);

	stCurrentTime.wHour = 0;
	stCurrentTime.wMinute = 0;
	stCurrentTime.wSecond = 0;
	stCurrentTime.wMilliseconds = 0;

	SystemTimeToFileTime(&stCurrentTime, &ftCurrentTime);


	return ftCurrentTime;
}

FILETIME calcEnd() {
	SYSTEMTIME stCurrentTime;
	FILETIME ftCurrentTime;
	ULONGLONG qwResult;

	GetSystemTime(&stCurrentTime);
	SystemTimeToFileTime(&stCurrentTime, &ftCurrentTime);
	qwResult = (((ULONGLONG)ftCurrentTime.dwHighDateTime) << 32) + ftCurrentTime.dwLowDateTime;  //Put current system time into quad word

	qwResult -= 1 * _DAY;  //Yesterday

	ftCurrentTime.dwLowDateTime = (DWORD)(qwResult & 0xFFFFFFFF);
	ftCurrentTime.dwHighDateTime = (DWORD)(qwResult >> 32);
	FileTimeToSystemTime(&ftCurrentTime, &stCurrentTime);

	stCurrentTime.wHour = 23;
	stCurrentTime.wMinute = 59;
	stCurrentTime.wSecond = 59;
	stCurrentTime.wMilliseconds = 999;

	SystemTimeToFileTime(&stCurrentTime, &ftCurrentTime);

	return ftCurrentTime;
}

DWORD queryLog(SYSTEMTIME& start, SYSTEMTIME& end)
{
	LPCWSTR queryxpath = L"*[(System/EventID=4801 or System/EventID=4800) and EventData/Data[@Name='TargetUserName']='garre']";
	DWORD status = ERROR_SUCCESS;
	EVT_HANDLE hResults = NULL;
	LPCWSTR pwsPath = L"SECURITY";

	hResults = EvtQuery(NULL, pwsPath, queryxpath, EvtQueryChannelPath | EvtQueryReverseDirection);
	if (NULL == hResults) {  //Handle to query results 
		status = GetLastError();

		if (ERROR_EVT_CHANNEL_NOT_FOUND == status) {
			wprintf(L"The channel was not found\n");
		}
		else if (ERROR_EVT_INVALID_QUERY) {
			wprintf(L"The query is not valid");
		}
		else {
			wprintf(L"EvtQuery Failed with %lu.\n", status);
		}
	}

	else {
		wprintf(L"Success");
		GetResults(hResults, start, end);
		
	}

	if (hResults) {
		EvtClose(hResults);
	}

	return status;
}


void main(void)
{
	DWORD status;
	SYSTEMTIME start;
	SYSTEMTIME end;
	CLSID clsid;
	HRESULT hRes;
	//this can be found by using oleview from VC cmd prompt
	LCID lcid;
	Excel::_ApplicationPtr pXL;


	queryLog(start, end);
	wprintf(L"First unlock of yesterday: %d:%d:%d", start.wHour, start.wMinute, start.wSecond);
	wprintf(L"Last lock of yesterday: %d:%d:%d", end.wHour, end.wMinute, end.wSecond);

	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	pXL.CreateInstance("Excel.Application");
	pXL->Workbooks->Open(L"E:\\Book2.xlsx");
	pXL->Visible = true;
	

	Excel::_WorksheetPtr pWksheet = pXL->ActiveSheet;
	Excel::RangePtr pRange = pWksheet->Cells;

	pRange->Item[1][1] = 5.4321;

	//pRange->

	//pXL-PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_FALSE);

	// Save the values in book.xml and release resources
	//pWksheet->SaveAs("C:\\dump\\book.xls");
	//pWksheet->Release();

	// And switch back on again...
	//pXL->PutDisplayAlerts(LOCALE_USER_DEFAULT, VARIANT_TRUE);

	//pXL->Quit();
	/*
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

	hRes = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_Application, (void**)&pXlApp);
	lcid = GetUserDefaultLCID();
	SetVisible(pXlApp, lcid);*/
	
	
}

DWORD getEventInfo(EVT_HANDLE hEvent, FILETIME& ftTemp, DWORD& dweventID)
{
	DWORD status = ERROR_SUCCESS;
	EVT_HANDLE hContext = NULL;
	DWORD dwBufferSize = 0;
	DWORD dwBufferUsed = 0;
	DWORD dwPropertyCount = 0;
	PEVT_VARIANT pRenderedValues = NULL;
	WCHAR wsGuid[50];
	LPWSTR pwsSid = NULL;
	ULONGLONG ullTimeStamp = 0;
	ULONGLONG ullNanoseconds = 0;
	SYSTEMTIME st;
	FILETIME ft;

	hContext = EvtCreateRenderContext(0, NULL, EvtRenderContextSystem);
	if (NULL == hContext)
	{
		wprintf(L"Failed with %lu", status = GetLastError());
		goto cleanup;
	}

	if (!EvtRender(hContext, hEvent, EvtRenderEventValues, dwBufferSize, pRenderedValues, &dwBufferUsed, &dwPropertyCount))
	{
		if (ERROR_INSUFFICIENT_BUFFER == (status = GetLastError()))
		{
			dwBufferSize = dwBufferUsed;
			pRenderedValues = (PEVT_VARIANT)malloc(dwBufferSize);
			status = ERROR_SUCCESS;
			if (pRenderedValues)
			{
				EvtRender(hContext, hEvent, EvtRenderEventValues, dwBufferSize, pRenderedValues, &dwBufferUsed, &dwPropertyCount);
				ullTimeStamp = pRenderedValues[EvtSystemTimeCreated].FileTimeVal;
				dweventID = pRenderedValues[EvtSystemEventID].UInt16Val;
				ftTemp.dwHighDateTime = (DWORD)((ullTimeStamp >> 32) & 0xFFFFFFFF);
				ftTemp.dwLowDateTime = (DWORD)(ullTimeStamp & 0xFFFFFFFF);
			}
			else
			{
				wprintf(L"malloc failed\n");
				status = ERROR_OUTOFMEMORY;
				goto cleanup;
			}
		}
	}
cleanup:
	return status;
}

DWORD GetResults(EVT_HANDLE hResults, SYSTEMTIME& ststart, SYSTEMTIME& stend) {

	DWORD status = ERROR_SUCCESS;
	EVT_HANDLE hEvents[ARRAY_SIZE];
	DWORD dwReturned = 0;
	FILETIME lastlock;
	FILETIME firstunlock;
	FILETIME ftTemp;
	FILETIME ftlocal;
	FILETIME ftstart = calcStart();
	FILETIME ftend = calcEnd();
	SYSTEMTIME stlastlock;
	SYSTEMTIME stfirstunlock;
	lastlock = ftstart;
	firstunlock = ftend;
	DWORD dwEventID;
	SYSTEMTIME temp;

	while (EvtNext(hResults, ARRAY_SIZE, hEvents, INFINITE, 0, &dwReturned))
	{
		for (DWORD i = 0; i < dwReturned; i++)
		{
			//if (ERROR_SUCCESS == (status = PrintEvent(hEvents[i])))
			if (ERROR_SUCCESS == (status = getEventInfo(hEvents[i], ftTemp, dwEventID)))
			{
				FileTimeToLocalFileTime(&ftTemp, &ftlocal);  //Convert to localtime before compare
				//Compare time
				if ((-1 == CompareFileTime(&ftlocal, &ftend)) && (-1 == CompareFileTime(&ftstart, &ftlocal)))//Make sure the time is within yesterday
				{
					FileTimeToSystemTime(&ftlocal, &temp);
					wprintf(L"Found event\n");

					//Find earliest unlock
					if (dwEventID == 4801) {
						if (-1 == CompareFileTime(&ftlocal, &firstunlock)) {  //If event is before earliest seen unlock time
							firstunlock = ftlocal;
						}
					}
					//Find latest lock
					if (dwEventID == 4800) {
						if (1 == CompareFileTime(&ftlocal, &lastlock)) {
							lastlock = ftlocal;
						}
					}
				}
				EvtClose(hEvents[i]);
				hEvents[i] = NULL;
			}
		}
	}

	if (ERROR_NO_MORE_ITEMS != (status = GetLastError()))
	{
		wprintf(L"EvtNext failed with %lu\n", status);
	}

	FileTimeToSystemTime(&firstunlock, &ststart);
	FileTimeToSystemTime(&lastlock, &stend);

	for (DWORD i = 0; i < dwReturned; i++)
	{
		if (NULL != hEvents[i])
			EvtClose(hEvents[i]);
	}

	return status;

}
