/**
  BSD 3-Clause License
  Copyright (c) 2019, Odzhan. All rights reserved.
  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  * Redistributions of source code must retain the above copyright notice, this
    list of conditions and the following disclaimer.
  * Redistributions in binary form must reproduce the above copyright notice,
    this list of conditions and the following disclaimer in the documentation
    and/or other materials provided with the distribution.
  * Neither the name of the copyright holder nor the names of its
    contributors may be used to endorse or promote products derived from
    this software without specific prior written permission.
  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
  FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
  DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
  SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
  CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
  OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
  OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#define UNICODE

#include <windows.h>
#include <initguid.h>
#include <activscp.h>
#include <objbase.h>
#include <shlwapi.h>
#include <comcat.h>

#include <string>
#include <iostream>
#include <vector>

#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "advapi32.lib")

void GetProgIDInfo(
    OLECHAR* pszLanguage,
    OLECHAR path[MAX_PATH],
    OLECHAR desc[MAX_PATH])
{
    LONG    res;
    HKEY    hKey;
    DWORD   len;
    wchar_t subkey[MAX_PATH];

    _snwprintf(subkey, MAX_PATH, L"CLSID\\%s", pszLanguage);

    len = MAX_PATH;

    res = RegGetValue(
        HKEY_CLASSES_ROOT, subkey,
        NULL, RRF_RT_REG_SZ,
        NULL, desc, &len);

    res = RegOpenKeyEx(
        HKEY_CLASSES_ROOT, subkey,
        0, KEY_READ, &hKey);

    if (res == ERROR_SUCCESS) {
        len = MAX_PATH;

        res = RegGetValue(
            hKey, L"InprocServer32",
            NULL, RRF_RT_REG_SZ,
            NULL, path, &len);

        RegCloseKey(hKey);
    }
}

void DisplayScriptEngines(void) {
    ICatInformation* pci = NULL;
    IEnumCLSID* pec = NULL;
    HRESULT         hr;
    CLSID           clsid;
    OLECHAR* progID, * idStr, path[MAX_PATH], desc[MAX_PATH];

    // initialize COM
    CoInitialize(NULL);

    // obtain component category manager for this machine
    hr = CoCreateInstance(
        CLSID_StdComponentCategoriesMgr,
        0, CLSCTX_SERVER, IID_ICatInformation,
        (void**)&pci);

    if (hr == S_OK) {
        // obtain list of script engine parsers
        hr = pci->EnumClassesOfCategories(
            1, &CATID_ActiveScriptParse, 0, 0, &pec);

        if (hr == S_OK) {
            // print each CLSID and Program ID
            for (;;) {
                ZeroMemory(path, ARRAYSIZE(path));
                ZeroMemory(desc, ARRAYSIZE(desc));

                hr = pec->Next(1, &clsid, 0);
                if (hr != S_OK) {
                    break;
                }
                ProgIDFromCLSID(clsid, &progID);
                StringFromCLSID(clsid, &idStr);
                GetProgIDInfo(idStr, path, desc);

                wprintf(L"\n*************************************\n");
                wprintf(L"Description : %s\n", desc);
                wprintf(L"CLSID       : %s\n", idStr);
                wprintf(L"Program ID  : %s\n", progID);
                wprintf(L"Path of DLL : %s\n", path);

                CoTaskMemFree(progID);
                CoTaskMemFree(idStr);
            }
            pec->Release();
        }
        pci->Release();
    }
}

int main(void) {
    DisplayScriptEngines();
    return 0;
}