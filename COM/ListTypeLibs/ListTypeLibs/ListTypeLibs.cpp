// ListTypeLibs.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <windows.h>
#include <iostream>

#include "TypeLib.h"
#include "JSON.h"



std::wstring variant_to_type(const char* variant)
{
    if (strcmp(variant, "VT_VOID") == 0)
        return L"void";
    if (strcmp(variant, "VT_I4") == 0)
        return L"int";
    if (strcmp(variant, "VT_INT") == 0)
        return L"int";
    if (strcmp(variant, "VT_BOOL") == 0)
        return L"bool";
    if (strcmp(variant, "VT_BSTR") == 0)
        return L"string";
    if (strcmp(variant, "VT_PTR") == 0)
        return L"Object&";
    if (strcmp(variant, "VT_DISPATCH") == 0)
        return L"object";
    if (strcmp(variant, "VT_HRESULT") == 0)
        return L"HRESULT";
    if (strcmp(variant, "VT_VARIANT") == 0)
        return L"VARIANT";

    if (strcmp(variant, "VT_VOID*") == 0)
        return L"void*";
    if (strcmp(variant, "VT_I4*") == 0)
        return L"int*";
    if (strcmp(variant, "VT_INT*") == 0)
        return L"int*";
    if (strcmp(variant, "VT_BOOL*") == 0)
        return L"bool*";
    if (strcmp(variant, "VT_BSTR*") == 0)
        return L"string*";
    if (strcmp(variant, "VT_DISPATCH*") == 0)
        return L"object*";
    if (strcmp(variant, "VT_HRESULT*") == 0)
        return L"HRESULT*";
    if (strcmp(variant, "VT_VARIANT*") == 0)
        return L"VARIANT*";

    return std::wstring(s2ws(variant));
}

JSONValue* get_typelib_info(const char* typelib_path)
{
    JSONObject typelib_info;

    TypeLib t;
    if (!t.Open(typelib_path)) {
        printf("Could not open typelib: %s\n", typelib_path);
        return NULL;
    }

    typelib_info[L"LibDocumentation"] = new JSONValue(s2ws(t.LibDocumentation()));

    JSONArray type_info_array;
    while (t.NextTypeInfo()) {
        JSONObject type_info;

        type_info[L"Interface"] = new JSONValue(s2ws(t.TypeDocumentation()));

        std::wstring interface_type = L"Unknown";
        if (t.IsTypeEnum())   interface_type = L"Enum";
        if (t.IsTypeRecord())   interface_type = L"Record";
        if (t.IsTypeModule())   interface_type = L"Module";
        if (t.IsTypeInterface())   interface_type = L"Interface";
        if (t.IsTypeDispatch())   interface_type = L"Dispatch";
        if (t.IsTypeCoClass())   interface_type = L"CoClass";
        if (t.IsTypeAlias())   interface_type = L"Alias";
        if (t.IsTypeUnion())   interface_type = L"Union";
        if (t.IsTypeMax())   interface_type = L"Max";

        type_info[L"InterfaceType"] = new JSONValue(interface_type);

        // Functions
        int nofFunctions = t.NofFunctions();

        JSONArray function_array;
        while (t.NextFunction()) {
            JSONObject function;

            std::string function_name = t.FunctionName();

            if (function_name.compare("QueryInterface") == 0 || function_name.compare("AddRef") == 0 || function_name.compare("Release") == 0 || function_name.compare("GetTypeInfoCount") == 0 || function_name.compare("GetTypeInfo") == 0 || function_name.compare("GetIDsOfNames") == 0 || function_name.compare("Invoke") == 0)
                continue;

            function[L"Name"] = new JSONValue(s2ws(t.FunctionName()));
            function[L"ReturnType"] = new JSONValue(variant_to_type(t.ReturnType().c_str()));

            JSONArray flags;
            if (t.HasFunctionTypeFlag(TypeLib::FDEFAULT)) flags.push_back(new JSONValue(L"FDEFAULT"));
            if (t.HasFunctionTypeFlag(TypeLib::FSOURCE)) flags.push_back(new JSONValue(L"FSOURCE"));
            if (t.HasFunctionTypeFlag(TypeLib::FRESTRICTED)) flags.push_back(new JSONValue(L"FRESTRICTED"));
            if (t.HasFunctionTypeFlag(TypeLib::FDEFAULTVTABLE)) flags.push_back(new JSONValue(L"FDEFAULTVTABLE"));
            function[L"Flags"] = new JSONValue(flags);

            TypeLib::INVOKEKIND ik = t.InvokeKind();
            switch (ik) {
            case TypeLib::func:
                function[L"InvokeKind"] = new JSONValue(L"function");
                break;
            case TypeLib::put:
                function[L"InvokeKind"] = new JSONValue(L"put");
                break;
            case TypeLib::get:
                function[L"InvokeKind"] = new JSONValue(L"get");
                break;
            case TypeLib::putref:
                function[L"InvokeKind"] = new JSONValue(L"put ref");
                break;
            default:
                function[L"InvokeKind"] = new JSONValue(L"unknown");
            }

            JSONArray parameter_array;
            while (t.NextParameter()) {
                JSONObject parameter;

                parameter[L"Name"] = new JSONValue(s2ws(t.ParameterName()));
                parameter[L"Type"] = new JSONValue(variant_to_type(t.ParameterType().c_str()));

                JSONArray tags;
                if (t.ParameterIsIn()) tags.push_back(new JSONValue(L"in"));
                if (t.ParameterIsOut()) tags.push_back(new JSONValue(L"out"));
                if (t.ParameterIsFLCID()) tags.push_back(new JSONValue(L"flcid"));
                if (t.ParameterIsReturnValue()) tags.push_back(new JSONValue(L"ret"));
                if (t.ParameterIsOptional()) tags.push_back(new JSONValue(L"opt"));
                if (t.ParameterHasDefault()) tags.push_back(new JSONValue(L"def"));
                if (t.ParameterHasCustumData()) tags.push_back(new JSONValue(L"cust"));
                parameter[L"Tags"] = new JSONValue(tags);

                parameter_array.push_back(new JSONValue(parameter));
            }
            function[L"Parameters"] = new JSONValue(parameter_array);

            function_array.push_back(new JSONValue(function));
        }
        type_info[L"Functions"] = new JSONValue(function_array);


        // Variables
        int nofVariables = t.NofVariables();


        type_info_array.push_back(new JSONValue(type_info));
    }

    typelib_info[L"TypeInfo"] = new JSONValue(type_info_array);

    return new JSONValue(typelib_info);
}

JSONValue* get_typelib_info_by_clsid(const CLSID& clsid)
{
    JSONObject typelib_info;

    TypeLib t;
    if (!t.OpenCLSID(clsid)) {
        //printf("Could not open typelib\n", clsid);
        return NULL;
    }

    typelib_info[L"LibDocumentation"] = new JSONValue(s2ws(t.LibDocumentation()));

    JSONArray type_info_array;
    while (t.NextTypeInfo()) {
        JSONObject type_info;

        type_info[L"Interface"] = new JSONValue(s2ws(t.TypeDocumentation()));

        std::wstring interface_type = L"Unknown";
        if (t.IsTypeEnum())   interface_type = L"Enum";
        if (t.IsTypeRecord())   interface_type = L"Record";
        if (t.IsTypeModule())   interface_type = L"Module";
        if (t.IsTypeInterface())   interface_type = L"Interface";
        if (t.IsTypeDispatch())   interface_type = L"Dispatch";
        if (t.IsTypeCoClass())   interface_type = L"CoClass";
        if (t.IsTypeAlias())   interface_type = L"Alias";
        if (t.IsTypeUnion())   interface_type = L"Union";
        if (t.IsTypeMax())   interface_type = L"Max";

        type_info[L"InterfaceType"] = new JSONValue(interface_type);

        // Functions
        int nofFunctions = t.NofFunctions();

        JSONArray function_array;
        while (t.NextFunction()) {
            JSONObject function;

            std::string function_name = t.FunctionName();

            if (function_name.compare("QueryInterface") == 0 || function_name.compare("AddRef") == 0 || function_name.compare("Release") == 0 || function_name.compare("GetTypeInfoCount") == 0 || function_name.compare("GetTypeInfo") == 0 || function_name.compare("GetIDsOfNames") == 0 || function_name.compare("Invoke") == 0)
                continue;

            function[L"Name"] = new JSONValue(s2ws(t.FunctionName()));
            function[L"ReturnType"] = new JSONValue(variant_to_type(t.ReturnType().c_str()));

            JSONArray flags;
            if (t.HasFunctionTypeFlag(TypeLib::FDEFAULT)) flags.push_back(new JSONValue(L"FDEFAULT"));
            if (t.HasFunctionTypeFlag(TypeLib::FSOURCE)) flags.push_back(new JSONValue(L"FSOURCE"));
            if (t.HasFunctionTypeFlag(TypeLib::FRESTRICTED)) flags.push_back(new JSONValue(L"FRESTRICTED"));
            if (t.HasFunctionTypeFlag(TypeLib::FDEFAULTVTABLE)) flags.push_back(new JSONValue(L"FDEFAULTVTABLE"));
            function[L"Flags"] = new JSONValue(flags);

            TypeLib::INVOKEKIND ik = t.InvokeKind();
            switch (ik) {
            case TypeLib::func:
                function[L"InvokeKind"] = new JSONValue(L"function");
                break;
            case TypeLib::put:
                function[L"InvokeKind"] = new JSONValue(L"put");
                break;
            case TypeLib::get:
                function[L"InvokeKind"] = new JSONValue(L"get");
                break;
            case TypeLib::putref:
                function[L"InvokeKind"] = new JSONValue(L"put ref");
                break;
            default:
                function[L"InvokeKind"] = new JSONValue(L"unknown");
            }

            JSONArray parameter_array;
            while (t.NextParameter()) {
                JSONObject parameter;

                parameter[L"Name"] = new JSONValue(s2ws(t.ParameterName()));
                parameter[L"Type"] = new JSONValue(variant_to_type(t.ParameterType().c_str()));

                JSONArray tags;
                if (t.ParameterIsIn()) tags.push_back(new JSONValue(L"in"));
                if (t.ParameterIsOut()) tags.push_back(new JSONValue(L"out"));
                if (t.ParameterIsFLCID()) tags.push_back(new JSONValue(L"flcid"));
                if (t.ParameterIsReturnValue()) tags.push_back(new JSONValue(L"ret"));
                if (t.ParameterIsOptional()) tags.push_back(new JSONValue(L"opt"));
                if (t.ParameterHasDefault()) tags.push_back(new JSONValue(L"def"));
                if (t.ParameterHasCustumData()) tags.push_back(new JSONValue(L"cust"));
                parameter[L"Tags"] = new JSONValue(tags);

                parameter_array.push_back(new JSONValue(parameter));
            }
            function[L"Parameters"] = new JSONValue(parameter_array);

            function_array.push_back(new JSONValue(function));
        }
        type_info[L"Functions"] = new JSONValue(function_array);


        // Variables
        int nofVariables = t.NofVariables();


        type_info_array.push_back(new JSONValue(type_info));
    }

    typelib_info[L"TypeInfo"] = new JSONValue(type_info_array);

    return new JSONValue(typelib_info);
}

std::wstring PrintCOMObjectName(const std::wstring& clsidKeyName) {
    HKEY hKey;
    std::wstring classKey = L"CLSID\\" + clsidKeyName;
    std::wstring comObjectName;

    if (RegOpenKeyEx(HKEY_CLASSES_ROOT, classKey.c_str(), 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        WCHAR name[256];
        DWORD nameSize = sizeof(name);
        if (RegQueryValueEx(hKey, nullptr, nullptr, nullptr, (LPBYTE)name, &nameSize) == ERROR_SUCCESS) {
            comObjectName = name;
        }
        RegCloseKey(hKey);
    }

    return comObjectName;
}

void PrintTypeLibInfo(const CLSID& clsid) {
    ITypeLib* pTypeLib = nullptr;
    HRESULT hr = LoadRegTypeLib(clsid, 1, 0, LOCALE_USER_DEFAULT, &pTypeLib);
    if (SUCCEEDED(hr)) {
        BSTR bstrFileName;
        hr = pTypeLib->GetDocumentation(-1, nullptr, &bstrFileName, nullptr, nullptr);
        if (SUCCEEDED(hr)) {
            wprintf(L"  TypeLib Path: %s\n", bstrFileName);
            SysFreeString(bstrFileName);
        }
        pTypeLib->Release();
    }
}

std::wstring GetTypeLibPathFromRegistry(const std::wstring& typelibGUID, const std::wstring& version) {
    HKEY hKey;
    std::wstring registryPath = L"SOFTWARE\\Classes\\TypeLib\\" + typelibGUID + L"\\" + version + L"\\0\\win64";
    std::wstring typelibPath;

    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, registryPath.c_str(), 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        wchar_t buffer[512];
        DWORD bufferSize = sizeof(buffer);
        if (RegQueryValueExW(hKey, NULL, NULL, NULL, (LPBYTE)buffer, &bufferSize) == ERROR_SUCCESS) {
            typelibPath = buffer;
        }
        RegCloseKey(hKey);
    }
    return typelibPath;
}

JSONValue* list_CLSIDs() 
{
    JSONArray clsid_info_array;

    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        DWORD index = 0;
        WCHAR clsidKeyName[256];
        DWORD clsidKeyNameSize = sizeof(clsidKeyName) / sizeof(clsidKeyName[0]);

        while (RegEnumKeyEx(hKey, index, clsidKeyName, &clsidKeyNameSize, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
            CLSID clsid;
            if (CLSIDFromString(clsidKeyName, &clsid) == S_OK) {
                //wprintf(L"Found CLSID: %s\n", clsidKeyName);

                JSONObject clsid_info;

                clsid_info[L"CLSID"] = new JSONValue(clsidKeyName);

                std::wstring comObjectName = PrintCOMObjectName(clsidKeyName);
                if (!comObjectName.empty()) {

                    clsid_info[L"Name"] = new JSONValue(comObjectName);
                    //wprintf(L"  COM Object Name: %s\n", comObjectName.c_str());
                }
                else {
                    //wprintf(L"  COM Object Name: Not found\n");
                }

                // Open the CLSID subkey
                HKEY hSubKey;
                std::wstring subKeyPath = L"CLSID\\" + std::wstring(clsidKeyName) + L"\\TypeLib";
                if (RegOpenKeyExW(HKEY_CLASSES_ROOT, subKeyPath.c_str(), 0, KEY_READ, &hSubKey) == ERROR_SUCCESS) {
                    // Query the TypeLib GUID
                    wchar_t typelibGUID[512];
                    DWORD typelibGUIDSize = sizeof(typelibGUID);
                    if (RegQueryValueExW(hSubKey, NULL, NULL, NULL, (LPBYTE)typelibGUID, &typelibGUIDSize) == ERROR_SUCCESS) {
                        //std::wcout << L"CLSID: " << clsid << L" has TypeLib GUID: " << typelibGUID << std::endl;

                        // Query the version key
                        std::wstring versionKeyPath = L"CLSID\\" + std::wstring(clsidKeyName) + L"\\Version";
                        wchar_t version[50];
                        DWORD versionSize = sizeof(version);
                        if (RegQueryValueExW(hSubKey, versionKeyPath.c_str(), NULL, NULL, (LPBYTE)version, &versionSize) != ERROR_SUCCESS) {
                            // Default to version 1.0 if no specific version is found
                            wcscpy_s(version, L"1.0");
                        }

                        // Get the TypeLib path for the specified GUID and version
                        std::wstring typelibPath = GetTypeLibPathFromRegistry(typelibGUID, version);
                        if (!typelibPath.empty()) {
                            JSONValue* typelib_info = get_typelib_info(ws2s(typelibPath).c_str()); // Call hello with the typelib path

                            if (typelib_info != NULL)
                            {
                                clsid_info[L"TypeLib"] = typelib_info;

                                clsid_info_array.push_back(new JSONValue(clsid_info));
                            }
                        }
                        else {
                            //std::wcerr << L"Failed to retrieve TypeLib path for CLSID: " << clsid << std::endl;
                        }
                    }

                    RegCloseKey(hSubKey);
                }

                /*

                PrintTypeLibInfo(clsid);
                JSONValue* typelib_info = get_typelib_info_by_clsid(clsid);

                //PrintTypeLibInfo(clsid);

                if (typelib_info != NULL)
                {
                    clsid_info[L"TypeLib"] = typelib_info;

                    clsid_info_array.push_back(new JSONValue(clsid_info));
                }
                */
                    
            }
            clsidKeyNameSize = sizeof(clsidKeyName) / sizeof(clsidKeyName[0]);
            index++;
        }
        RegCloseKey(hKey);
    }
    else {
        std::cerr << "Failed to open CLSID registry key." << std::endl;

        return NULL;
    }

    return new JSONValue(clsid_info_array);
}

int main(int argc, char* argv[]) 
{
    if (argc == 2) {
        CoInitialize(nullptr); // Initialize COM library

        JSONValue* clsid_info = list_CLSIDs();

        std::cout << "First argument: " << argv[1] << std::endl;

        wprintf(L"> %s\n", clsid_info->Stringify().c_str());

        CoUninitialize(); // Uninitialize COM library

        // Create or open the file
        HANDLE hFile = CreateFileA(
            argv[1],                   // File name
            GENERIC_WRITE,              // Desired access (write)
            0,                          // Share mode (none)
            NULL,                       // Security attributes (default)
            CREATE_ALWAYS,              // Create a new file or overwrite if it exists
            FILE_ATTRIBUTE_NORMAL,      // File attributes (normal)
            NULL                        // Template file (none)
        );

        // Check if file handle is valid
        if (hFile == INVALID_HANDLE_VALUE) {
            std::cerr << "Could not create or open file, error: " << GetLastError() << std::endl;
            return 1;
        }

        std::string json = ws2s(clsid_info->Stringify());

        DWORD bytesWritten = 0;
        // Write the data to the file
        BOOL result = WriteFile(
            hFile,                      // File handle
            json.c_str(),                       // Data to write
            json.length(),               // Number of bytes to write
            &bytesWritten,              // Number of bytes written
            NULL                        // Overlapped (not used)
        );

        // Check if the write operation succeeded
        if (!result) {
            std::cerr << "Failed to write to file, error: " << GetLastError() << std::endl;
            CloseHandle(hFile);
            return 1;
        }
        else {
            std::cout << "Successfully wrote " << bytesWritten << " bytes to the file." << std::endl;
        }

        // Close the file handle
        CloseHandle(hFile);

    }
    else {
        std::cout << "No arguments provided!\n Usage ListTypeLibs.exe <output_file>\n" << std::endl;
    }

    return 0;
}


int test()
{
    TypeLib t;
    if (!t.Open("C:\\Windows\\System32\\wshom.ocx")) {
        std::cout << "Couldn't open type library" << std::endl;
        return -1;
    }

    std::cout << t.LibDocumentation() << std::endl;

    int nofTypeInfos = t.NofTypeInfos();


    std::cout << "Nof Type Infos: " << nofTypeInfos << std::endl;

    while (t.NextTypeInfo()) {
        std::string type_doc = t.TypeDocumentation();

        std::cout << std::endl;
        std::cout << type_doc << std::endl;
        std::cout << "----------------------------" << std::endl;

        std::cout << "  Interface: ";
        if (t.IsTypeEnum())   std::cout << "Enum";
        if (t.IsTypeRecord())   std::cout << "Record";
        if (t.IsTypeModule())   std::cout << "Module";
        if (t.IsTypeInterface())   std::cout << "Interface";
        if (t.IsTypeDispatch())   std::cout << "Dispatch";
        if (t.IsTypeCoClass())   std::cout << "CoClass";
        if (t.IsTypeAlias())   std::cout << "Alias";
        if (t.IsTypeUnion())   std::cout << "Union";
        if (t.IsTypeMax())   std::cout << "Max";
        std::cout << std::endl;

        int nofFunctions = t.NofFunctions();
        int nofVariables = t.NofVariables();

        std::cout << "  functions: " << nofFunctions << std::endl;
        std::cout << "  variables: " << nofVariables << std::endl;

        while (t.NextFunction()) {
            std::cout << std::endl;
            std::cout << "  Function     : " << t.FunctionName() << std::endl;

            std::cout << "    returns    : " << t.ReturnType() << std::endl;

            std::cout << "    flags      : ";
            if (t.HasFunctionTypeFlag(TypeLib::FDEFAULT)) std::cout << "FDEFAULT ";
            if (t.HasFunctionTypeFlag(TypeLib::FSOURCE)) std::cout << "FSOURCE ";
            if (t.HasFunctionTypeFlag(TypeLib::FRESTRICTED)) std::cout << "FRESTRICTED ";
            if (t.HasFunctionTypeFlag(TypeLib::FDEFAULTVTABLE)) std::cout << "FDEFAULTVTABLE ";
            std::cout << std::endl;

            TypeLib::INVOKEKIND ik = t.InvokeKind();
            switch (ik) {
            case TypeLib::func:
                std::cout << "    invoke kind: function" << std::endl;
                break;
            case TypeLib::put:
                std::cout << "    invoke kind: put" << std::endl;
                break;
            case TypeLib::get:
                std::cout << "    invoke kind: get" << std::endl;
                break;
            case TypeLib::putref:
                std::cout << "    invoke kind: put ref" << std::endl;
                break;
            default:
                std::cout << "    invoke kind: ???" << std::endl;
            }

            std::cout << "    params     : " << t.NofParameters() << std::endl;
            std::cout << "    params opt : " << t.NofOptionalParameters() << std::endl;

            while (t.NextParameter()) {
                std::cout << "    Parameter  : " << t.ParameterName();
                std::cout << " type = " << t.ParameterType();
                if (t.ParameterIsIn()) std::cout << " in";
                if (t.ParameterIsOut()) std::cout << " out";
                if (t.ParameterIsFLCID()) std::cout << " flcid";
                if (t.ParameterIsReturnValue()) std::cout << " ret";
                if (t.ParameterIsOptional()) std::cout << " opt";
                if (t.ParameterHasDefault()) std::cout << " def";
                if (t.ParameterHasCustumData()) std::cout << " cust";
                std::cout << std::endl;
            }
        }

        while (t.NextVariable()) {
            std::cout << "  Variable : " << t.VariableName() << std::endl;
            std::cout << "      Type : " << t.VariableType();
            TypeLib::VARIABLEKIND vk = t.VariableKind();
            switch (vk) {
            case TypeLib::instance: std::cout << " (instance)" << std::endl; break;
            case TypeLib::static_: std::cout << " (static)" << std::endl; break;
            case TypeLib::const_: std::cout << " (const ";
                std::cout << t.ConstValue() << ")" << std::endl;            break;
            case TypeLib::dispatch: std::cout << " (dispatch)" << std::endl; break;
            default:
                std::cout << "    variable kind: unknown" << std::endl;
            }
        }
    }

    return 0;
}
