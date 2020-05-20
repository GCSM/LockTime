#include <windows.h>
#include <stdio.h>
#include <winevt.h>

#pragma comment(lib, "wevtapi.lib")

void main(void)
{
    EVT_HANDLE hProviders = NULL;
    LPWSTR pwcsProviderName = NULL;
    LPWSTR pTemp = NULL;
    DWORD dwBufferSize = 0;
    DWORD dwBufferUsed = 0;
    DWORD status = ERROR_SUCCESS;

    // Get a handle to the list of providers.
    hProviders = EvtOpenPublisherEnum(NULL, 0);
    if (NULL == hProviders)
    {
        wprintf(L"EvtOpenPublisherEnum failed with %lu\n", GetLastError());
        goto cleanup;
    }

    wprintf(L"List of registered providers:\n\n");

    // Enumerate the providers in the list.
    while (true)
    {
        // Get a provider from the list. If the buffer is not big enough
        // to contain the provider's name, reallocate the buffer to the required size.
        if (!EvtNextPublisherId(hProviders, dwBufferSize, pwcsProviderName, &dwBufferUsed))
        {
            status = GetLastError();
            if (ERROR_NO_MORE_ITEMS == status)
            {
                break;
            }
            else if (ERROR_INSUFFICIENT_BUFFER == status)
            {
                dwBufferSize = dwBufferUsed;
                pTemp = (LPWSTR)realloc(pwcsProviderName, dwBufferSize * sizeof(WCHAR));
                if (pTemp)
                {
                    pwcsProviderName = pTemp;
                    pTemp = NULL;
                    EvtNextPublisherId(hProviders, dwBufferSize, pwcsProviderName, &dwBufferUsed);
                }
                else
                {
                    wprintf(L"realloc failed\n");
                    goto cleanup;
                }
            }

            if (ERROR_SUCCESS != (status = GetLastError()))
            {
                wprintf(L"EvtNextPublisherId failed with %d\n", status);
                goto cleanup;
            }
        }

        wprintf(L"%s\n", pwcsProviderName);

        RtlZeroMemory(pwcsProviderName, dwBufferUsed * sizeof(WCHAR));
    }

cleanup:

    if (pwcsProviderName)
        free(pwcsProviderName);

    if (hProviders)
        EvtClose(hProviders);
}