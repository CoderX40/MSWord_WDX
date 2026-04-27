#include "plugin_shared.h"

#include <string>

#ifndef _WIN64
#pragma comment(linker, "/export:ContentGetSupportedField=_ContentGetSupportedField@16")
#pragma comment(linker, "/export:ContentGetValue=_ContentGetValue@24")
#pragma comment(linker, "/export:ContentGetValueW=_ContentGetValueW@24")
#pragma comment(linker, "/export:ContentGetSupportedFieldFlags=_ContentGetSupportedFieldFlags@4")
#pragma comment(linker, "/export:ContentGetDefaultSortOrder=_ContentGetDefaultSortOrder@4")
#pragma comment(linker, "/export:ContentSetValue=_ContentSetValue@24")
#pragma comment(linker, "/export:ContentSetValueW=_ContentSetValueW@24")
#pragma comment(linker, "/export:ContentEditValue=_ContentEditValue@32")
#pragma comment(linker, "/export:ContentEditValueW=_ContentEditValueW@32")
#endif

namespace {

struct ContentDefaultParamStruct {
    int size;
    DWORD PluginInterfaceVersionLow;
    DWORD PluginInterfaceVersionHi;
    char DefaultIniName[MAX_PATH];
};

const char* SafeGetIniPathAnsi(void* dparm)
{
    if (!dparm) return nullptr;
    const ContentDefaultParamStruct* dps = static_cast<const ContentDefaultParamStruct*>(dparm);
    if (dps->size < static_cast<int>(offsetof(ContentDefaultParamStruct, DefaultIniName) + 1)) return nullptr;
    return dps->DefaultIniName[0] ? dps->DefaultIniName : nullptr;
}

} // namespace

extern "C" {

__declspec(dllexport) int __stdcall ContentGetSupportedField(int fieldIndex, char* fieldName, char* units, int maxLen)
{
    const FieldDescriptor* descriptor = GetFieldDescriptor(fieldIndex);
    if (!descriptor) return ft_nomorefields;

    strncpy_s(fieldName, maxLen, descriptor->name, _TRUNCATE);
    strncpy_s(units, maxLen, descriptor->units, _TRUNCATE);
    return descriptor->type;
}

__declspec(dllexport) int __stdcall ContentGetValueW(WCHAR* fileName, int fieldIndex, int unitIndex, void* fieldValue, int maxLen, int flags)
{
    UNREFERENCED_PARAMETER(flags);
    if (!fileName) return ft_fieldempty;

    std::string ansiPath;
    if (!WidePathToAnsi(fileName, ansiPath)) return ft_fieldempty;

    g_cancelRequested.store(false, std::memory_order_relaxed);
    return GetContentValueWInternal(ansiPath, fieldIndex, unitIndex, fieldValue, maxLen);
}

__declspec(dllexport) int __stdcall ContentGetValue(char* fileName, int fieldIndex, int unitIndex, void* fieldValue, int maxLen, int flags)
{
    UNREFERENCED_PARAMETER(flags);
    if (!fileName) return ft_fieldempty;

    g_cancelRequested.store(false, std::memory_order_relaxed);
    return GetContentValueInternal(fileName, fieldIndex, unitIndex, fieldValue, maxLen);
}

__declspec(dllexport) int __stdcall ContentGetDetectString(char* DetectString, int maxlen)
{
    if (!DetectString || maxlen <= 0) return 1;
    const char* shortOnly = "EXT=docx";
    if (static_cast<size_t>(maxlen) < strlen(shortOnly) + 1) return 1;
    strcpy_s(DetectString, static_cast<size_t>(maxlen), shortOnly);
    return 0;
}

__declspec(dllexport) int __stdcall ContentGetDetectStringW(WCHAR* DetectString, int maxlen)
{
    if (!DetectString || maxlen <= 0) return 1;
    const std::wstring shortOnly = L"EXT=docx";
    if (static_cast<size_t>(maxlen) < shortOnly.length() + 1) return 1;
    wcscpy_s(DetectString, static_cast<size_t>(maxlen), shortOnly.c_str());
    return 0;
}

__declspec(dllexport) void __stdcall ContentSetDefaultParams(void* dparm)
{
    if (!dparm) return;
    const char* pIniPath = SafeGetIniPathAnsi(dparm);
    if (pIniPath) SetPluginDefaultIniPath(pIniPath);
}

__declspec(dllexport) void __stdcall ContentStopGetValue(void)
{
    g_cancelRequested.store(true, std::memory_order_relaxed);
    ClearCache();
}

__declspec(dllexport) void __stdcall ContentStopGetValueW(void)
{
    ContentStopGetValue();
}

__declspec(dllexport) int __stdcall ContentGetSupportedFieldFlags(int fieldIndex)
{
    DbgLog("ContentGetSupportedFieldFlags called: fieldIndex=%d\n", fieldIndex);
    if (fieldIndex == -1) return contflags_edit | contflags_fieldedit;

    const FieldDescriptor* descriptor = GetFieldDescriptor(fieldIndex);
    return descriptor ? descriptor->flags : 0;
}

__declspec(dllexport) int __stdcall ContentGetDefaultSortOrder(int fieldIndex)
{
    switch (fieldIndex) {
    case FIELD_CORE_CREATED_DATE:
    case FIELD_CORE_MODIFIED_DATE:
    case FIELD_CORE_LAST_PRINTED_DATE:
    case FIELD_CORE_REVISION_NUMBER:
    case FIELD_APP_EDITING_TIME:
    case FIELD_APP_PAGES:
    case FIELD_APP_PARAGRAPHS:
    case FIELD_APP_LINES:
    case FIELD_APP_WORDS:
    case FIELD_APP_CHARACTERS:
    case FIELD_COMMENTS:
    case FIELD_TOTAL_REVISIONS:
    case FIELD_TOTAL_INSERTIONS:
    case FIELD_TOTAL_DELETIONS:
    case FIELD_TOTAL_MOVES:
    case FIELD_TOTAL_FORMATTING_CHANGES:
        return -1;
    default:
        return 1;
    }
}

__declspec(dllexport) void __stdcall ContentPluginUnloading(void)
{
    g_cancelRequested.store(true, std::memory_order_relaxed);
    ClearCache();
}

__declspec(dllexport) int __stdcall ContentEditValueW(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int maxLen, int flags, const char* langIdentifier)
{
    return RunContentEditValueW(parentWin, fieldIndex, unitIndex, fieldType, fieldValue, maxLen, flags, langIdentifier);
}

__declspec(dllexport) int __stdcall ContentEditValue(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, char* fieldValue, int maxLen, int flags, const char* langIdentifier)
{
    return RunContentEditValue(parentWin, fieldIndex, unitIndex, fieldType, fieldValue, maxLen, flags, langIdentifier);
}

__declspec(dllexport) int __stdcall ContentSetValueW(WCHAR* fileName, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int flags)
{
    return RunContentSetValueW(fileName, fieldIndex, unitIndex, fieldType, fieldValue, flags);
}

} // extern "C"

extern "C" __declspec(dllexport) int __stdcall ContentSetValue(char* fileName, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int flags)
{
    return RunContentSetValue(fileName, fieldIndex, unitIndex, fieldType, fieldValue, flags);
}
