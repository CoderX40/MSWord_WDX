#include "pch.h"

#include "plugin_shared.h"

#include <algorithm>
#include <cctype>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <functional>
#include <sstream>
#include <vector>

bool HasHiddenTextInDocumentXmlContent(const std::string& xmlContent);
void PopulateDerivedCachedParts(CachedParts& parts);

namespace {

std::mutex g_cacheMutex;
CachedParts g_cachedParts;
std::mutex g_languageMutex;
std::string g_defaultIniPath;
std::string g_cachedLanguageCode = "eng";
FILETIME g_cachedIniWriteTime{};

const FieldDescriptor kFieldDescriptors[FIELD_COUNT] = {
    {"Title", "", ft_stringw, contflags_edit},
    {"Subject", "", ft_stringw, contflags_edit},
    {"Author", "", ft_stringw, contflags_edit},
    {"Manager", "", ft_stringw, contflags_edit},
    {"Company", "", ft_stringw, contflags_edit},
    {"Keywords", "", ft_stringw, contflags_edit},
    {"Comments", "", ft_stringw, contflags_edit},
    {"Hyperlink base", "", ft_stringw, contflags_edit},
    {"Template", "", ft_stringw, 0},
    {"Created", "", ft_datetime, contflags_edit},
    {"Modified", "", ft_datetime, contflags_edit},
    {"Printed", "", ft_datetime, contflags_edit},
    {"Last saved by", "", ft_stringw, contflags_edit},
    {"Revision number", "", ft_numeric_32, contflags_edit},
    {"Total editing time", "min", ft_numeric_32, contflags_edit},
    {"Pages", "", ft_numeric_32, 0},
    {"Paragraphs", "", ft_numeric_32, 0},
    {"Lines", "", ft_numeric_32, 0},
    {"Words", "", ft_numeric_32, 0},
    {"Characters", "", ft_numeric_32, 0},
    {"Compatibility mode", "", ft_boolean, 0},
    {"Hidden text", "", ft_boolean, 0},
    {"Number of comments", "", ft_numeric_32, 0},
    {"Document protection", "", ft_stringw, contflags_edit | contflags_fieldedit},
    {"Auto update styles", "Yes|No", ft_multiplechoice, contflags_edit},
    {"Anonymisation", "No|Personal information|Date and time|Personal information, Date and time", ft_multiplechoice, contflags_edit},
    {"Contains tracked changes", "", ft_boolean, 0},
    {"Track changes mode", "Enabled|Disabled", ft_multiplechoice, contflags_edit},
    {"Tracked changes authors", "", ft_stringw, contflags_edit | contflags_fieldedit},
    {"Total revisions", "", ft_numeric_32, 0},
    {"Total insertions", "", ft_numeric_32, 0},
    {"Total deletions", "", ft_numeric_32, 0},
    {"Total moves", "", ft_numeric_32, 0},
    {"Total formatting changes", "", ft_numeric_32, 0},
};

bool IsProbablyXml(const char* name)
{
    if (!name) return false;
    if (strstr(name, "docProps/") || strstr(name, "word/") || strcmp(name, "[Content_Types].xml") == 0) return true;
    size_t n = strlen(name);
    return n > 4 && strcmp(name + n - 4, ".xml") == 0;
}

bool IsValidUtf8(const std::string& str)
{
    const unsigned char* bytes = reinterpret_cast<const unsigned char*>(str.c_str());
    size_t len = str.size();
    size_t i = 0;
    while (i < len) {
        unsigned char c = bytes[i];
        if (c < 0x80) {
            ++i;
        }
        else if ((c >> 5) == 0x6) {
            if (i + 1 >= len || (bytes[i + 1] >> 6) != 0x2) return false;
            i += 2;
        }
        else if ((c >> 4) == 0xE) {
            if (i + 2 >= len || (bytes[i + 1] >> 6) != 0x2 || (bytes[i + 2] >> 6) != 0x2) return false;
            i += 3;
        }
        else if ((c >> 3) == 0x1E) {
            if (i + 3 >= len || (bytes[i + 1] >> 6) != 0x2 || (bytes[i + 2] >> 6) != 0x2 || (bytes[i + 3] >> 6) != 0x2) return false;
            i += 4;
        }
        else {
            return false;
        }
    }
    return true;
}

std::string AnsiToUtf8(const std::string& input)
{
    if (input.empty()) return std::string();
    int wlen = MultiByteToWideChar(CP_ACP, 0, input.data(), static_cast<int>(input.size()), nullptr, 0);
    if (wlen <= 0) return std::string();

    std::wstring wide(static_cast<size_t>(wlen), L'\0');
    if (MultiByteToWideChar(CP_ACP, 0, input.data(), static_cast<int>(input.size()), &wide[0], wlen) == 0) return std::string();

    int ulen = WideCharToMultiByte(CP_UTF8, 0, wide.data(), wlen, nullptr, 0, nullptr, nullptr);
    if (ulen <= 0) return std::string();

    std::string output(static_cast<size_t>(ulen), '\0');
    if (WideCharToMultiByte(CP_UTF8, 0, wide.data(), wlen, &output[0], ulen, nullptr, nullptr) == 0) return std::string();
    return output;
}

void ExtractAuthorsRecursive(tinyxml2::XMLElement* elem, std::set<std::string>& authors)
{
    if (!elem) return;

    const char* name = elem->Name();
    if (name) {
        if (strcmp(name, "w:ins") == 0 ||
            strcmp(name, "w:del") == 0 ||
            strcmp(name, "w:rPrChange") == 0 ||
            strcmp(name, "w:pPrChange") == 0 ||
            strcmp(name, "w:sectPrChange") == 0 ||
            strcmp(name, "w:tblPrChange") == 0 ||
            strcmp(name, "w:tblGridChange") == 0 ||
            strcmp(name, "w:trPrChange") == 0 ||
            strcmp(name, "w:tcPrChange") == 0 ||
            strcmp(name, "w:shd") == 0 ||
            strcmp(name, "w:border") == 0 ||
            strcmp(name, "w:jc") == 0 ||
            strcmp(name, "w:ind") == 0 ||
            strcmp(name, "w:spacing") == 0 ||
            strcmp(name, "w:numPr") == 0 ||
            strcmp(name, "w:tabs") == 0 ||
            strcmp(name, "w:altChunk") == 0 ||
            strcmp(name, "w:smartTagPr") == 0 ||
            strcmp(name, "w:customXmlPr") == 0 ||
            strcmp(name, "w:sdtPr") == 0 ||
            strcmp(name, "w:style") == 0 ||
            strcmp(name, "w:tblLook") == 0) {
            const char* author = elem->Attribute("w:author");
            if (author) authors.insert(author);
            const char* originalAuthor = elem->Attribute("w:originalAuthor");
            if (originalAuthor) authors.insert(originalAuthor);
        }
    }

    for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child; child = child->NextSiblingElement()) {
        ExtractAuthorsRecursive(child, authors);
    }
}

bool IsOnOffElementEnabled(tinyxml2::XMLElement* element)
{
    if (!element) return false;
    const char* val = element->Attribute("w:val");
    if (!val) return true;

    std::string s(val);
    for (char& ch : s) {
        ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));
    }

    return !(s == "0" || s == "false" || s == "off");
}

void CountTrackedChangesRecursive(tinyxml2::XMLElement* elem, TrackedChangeCounts& counts)
{
    if (!elem || IsCanceled()) return;

    const char* name = elem->Name();
    if (name) {
        if (strcmp(name, "w:ins") == 0) {
            ++counts.insertions;
        }
        else if (strcmp(name, "w:del") == 0) {
            ++counts.deletions;
        }
        else if (strcmp(name, "w:moveFrom") == 0) {
            ++counts.moves;
        }
        else {
            std::string changeId = name;
            bool isFormattingChange =
                strcmp(name, "w:rPrChange") == 0 ||
                strcmp(name, "w:pPrChange") == 0 ||
                strcmp(name, "w:sectPrChange") == 0 ||
                strcmp(name, "w:tblPrChange") == 0 ||
                strcmp(name, "w:tblGridChange") == 0 ||
                strcmp(name, "w:trPrChange") == 0 ||
                strcmp(name, "w:tcPrChange") == 0;

            if (isFormattingChange) {
                if (elem->FirstChildElement()) changeId += ":" + std::string(elem->FirstChildElement()->Name());
                else changeId += ":noChild";

                if (counts.uniqueFormattingChanges.insert(changeId).second) {
                    ++counts.formattingChanges;
                }
            }
        }
    }

    for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child; child = child->NextSiblingElement()) {
        CountTrackedChangesRecursive(child, counts);
    }
}

int GetChoiceIndex(const void* fieldValue)
{
    if (fieldValue == nullptr || reinterpret_cast<uintptr_t>(fieldValue) < 0xFFFF) {
        if (fieldValue == nullptr) return -1;
        return static_cast<int>(reinterpret_cast<uintptr_t>(fieldValue));
    }

    return -1;
}

int MapChoiceTextToIndex(int fieldIndex, const char* text)
{
    if (!text || !*text) return -1;

    if (fieldIndex == FIELD_AUTO_UPDATE_STYLES) {
        if (strcmp(text, "Yes") == 0) return 0;
        if (strcmp(text, "No") == 0) return 1;
    }
    else if (fieldIndex == FIELD_TRACK_CHANGES_ENABLED_DISABLED) {
        if (strcmp(text, "Enabled") == 0) return 0;
        if (strcmp(text, "Disabled") == 0) return 1;
    }
    else if (fieldIndex == FIELD_ANONYMISATION) {
        if (strcmp(text, "No") == 0) return 0;
        if (strcmp(text, "Personal information") == 0) return 1;
        if (strcmp(text, "Date and time") == 0) return 2;
        if (strcmp(text, "Personal information, Date and time") == 0) return 3;
    }

    return -1;
}

std::wstring GetPluginLngPath()
{
    HMODULE module = nullptr;
    if (!GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCWSTR>(&GetPluginLngPath),
        &module)) {
        return std::wstring();
    }

    wchar_t path[MAX_PATH] = {};
    if (GetModuleFileNameW(module, path, MAX_PATH) == 0) {
        return std::wstring();
    }

    std::wstring result = path;
    size_t dot = result.find_last_of(L'.');
    if (dot == std::wstring::npos) result += L".lng";
    else {
        result.erase(dot);
        result += L".lng";
    }
    return result;
}

std::string ExtractLanguageCode(const std::string& value)
{
    if (value.empty()) return std::string();

    std::string trimmed = value;
    while (!trimmed.empty() && (trimmed.back() == '\r' || trimmed.back() == '\n' || trimmed.back() == ' ' || trimmed.back() == '\t')) {
        trimmed.pop_back();
    }

    size_t slash = trimmed.find_last_of("\\/");
    std::string tail = slash == std::string::npos ? trimmed : trimmed.substr(slash + 1);
    size_t dot = tail.find_last_of('.');
    if (dot != std::string::npos) tail.erase(dot);
    size_t underscore = tail.find_last_of('_');
    if (underscore != std::string::npos && underscore + 1 < tail.size()) tail = tail.substr(underscore + 1);

    if (tail.size() >= 3 && tail.size() <= 8) {
        for (char& ch : tail) ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));
        return tail;
    }

    return std::string();
}

std::string NormalizeLanguageCode(const std::string& rawCode)
{
    if (rawCode.empty()) return "eng";

    std::string code = rawCode;
    for (char& ch : code) {
        ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));
    }

    if (code == "en" || code == "eng" || code == "enu") return "eng";
    if (code == "fr" || code == "fre" || code == "fra") return "fra";
    if (code == "pt" || code == "ptg" || code == "por") return "por";

    return code;
}

std::vector<std::string> GetLanguageSectionCandidates(const std::string& rawCode)
{
    std::vector<std::string> candidates;

    auto addCandidate = [&](const std::string& value) {
        if (value.empty()) return;
        if (std::find(candidates.begin(), candidates.end(), value) == candidates.end()) {
            candidates.push_back(value);
        }
    };

    std::string normalized = NormalizeLanguageCode(rawCode);
    addCandidate(normalized);
    addCandidate(rawCode);

    if (normalized == "eng") addCandidate("enu");
    if (normalized == "enu") addCandidate("eng");

    addCandidate("eng");
    return candidates;
}

std::string GetDefaultIniPathFallback()
{
    const char* envNames[] = { "APPDATA", "LOCALAPPDATA", "WINDIR" };
    const char* suffixes[] = {
        "\\GHISLER\\wincmd.ini",
        "\\wincmd.ini"
    };

    char buffer[MAX_PATH] = {};
    for (const char* envName : envNames) {
        DWORD len = GetEnvironmentVariableA(envName, buffer, static_cast<DWORD>(std::size(buffer)));
        if (len == 0 || len >= std::size(buffer)) {
            continue;
        }

        for (const char* suffix : suffixes) {
            std::string candidate = std::string(buffer) + suffix;
            DWORD attrs = GetFileAttributesA(candidate.c_str());
            if (attrs != INVALID_FILE_ATTRIBUTES && !(attrs & FILE_ATTRIBUTE_DIRECTORY)) {
                return candidate;
            }
        }
    }

    return std::string();
}

std::string ResolveWinCmdIniPath(const std::string& iniPath)
{
    if (iniPath.empty()) return std::string();

    DWORD attrs = GetFileAttributesA(iniPath.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) {
        return std::string();
    }

    std::string lowered = iniPath;
    for (char& ch : lowered) {
        ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));
    }

    const std::string contplugName = "contplug.ini";
    const std::string wincmdName = "wincmd.ini";

    if (lowered.size() >= contplugName.size() &&
        lowered.compare(lowered.size() - contplugName.size(), contplugName.size(), contplugName) == 0) {
        std::string candidate = iniPath.substr(0, iniPath.size() - contplugName.size()) + wincmdName;
        DWORD candidateAttrs = GetFileAttributesA(candidate.c_str());
        if (candidateAttrs != INVALID_FILE_ATTRIBUTES && !(candidateAttrs & FILE_ATTRIBUTE_DIRECTORY)) {
            DbgLog("ResolveWinCmdIniPath: '%s' -> sibling '%s'\n", iniPath.c_str(), candidate.c_str());
            return candidate;
        }

        std::string fallback = GetDefaultIniPathFallback();
        if (!fallback.empty()) {
            DbgLog("ResolveWinCmdIniPath: sibling missing for '%s', fallback '%s'\n", iniPath.c_str(), fallback.c_str());
            return fallback;
        }
    }

    if (lowered.size() >= wincmdName.size() &&
        lowered.compare(lowered.size() - wincmdName.size(), wincmdName.size(), wincmdName) == 0) {
        return iniPath;
    }

    size_t slash = iniPath.find_last_of("\\/");
    if (slash != std::string::npos) {
        std::string candidate = iniPath.substr(0, slash + 1) + wincmdName;
        DWORD candidateAttrs = GetFileAttributesA(candidate.c_str());
        if (candidateAttrs != INVALID_FILE_ATTRIBUTES && !(candidateAttrs & FILE_ATTRIBUTE_DIRECTORY)) {
            DbgLog("ResolveWinCmdIniPath: '%s' -> directory candidate '%s'\n", iniPath.c_str(), candidate.c_str());
            return candidate;
        }
    }

    DbgLog("ResolveWinCmdIniPath: keeping '%s'\n", iniPath.c_str());
    return iniPath;
}

bool TryGetFileWriteTime(const std::string& path, FILETIME& writeTime)
{
    WIN32_FILE_ATTRIBUTE_DATA attrs{};
    if (!GetFileAttributesExA(path.c_str(), GetFileExInfoStandard, &attrs)) {
        return false;
    }

    writeTime = attrs.ftLastWriteTime;
    return true;
}

std::string DetectCurrentLanguageCodeFromIniPath(const std::string& iniPath)
{
    if (iniPath.empty()) return "eng";

    const char* sections[] = { "Configuration", "International", "TC" };
    const char* keys[] = { "LanguageINI", "LanguageIni", "Language", "Lang", "LanguageFile" };
    char buffer[MAX_PATH] = {};

    for (const char* section : sections) {
        for (const char* key : keys) {
            buffer[0] = '\0';
            GetPrivateProfileStringA(section, key, "", buffer, static_cast<DWORD>(sizeof(buffer)), iniPath.c_str());
            if (buffer[0]) {
                std::string code = ExtractLanguageCode(buffer);
                if (!code.empty()) return NormalizeLanguageCode(code);
            }
        }
    }

    return "eng";
}

std::string DetectCurrentLanguageCode()
{
    std::string iniPath = g_defaultIniPath;
    if (iniPath.empty()) {
        iniPath = GetDefaultIniPathFallback();
    }
    return DetectCurrentLanguageCodeFromIniPath(ResolveWinCmdIniPath(iniPath));
}

void RefreshPluginLanguageCacheIfNeeded()
{
    std::string iniPath;
    {
        std::lock_guard<std::mutex> cacheLock(g_languageMutex);
        iniPath = g_defaultIniPath;
    }

    if (iniPath.empty()) {
        iniPath = GetDefaultIniPathFallback();
    }
    iniPath = ResolveWinCmdIniPath(iniPath);
    if (iniPath.empty()) {
        return;
    }

    FILETIME currentWriteTime{};
    bool haveWriteTime = TryGetFileWriteTime(iniPath, currentWriteTime);

    {
        std::lock_guard<std::mutex> cacheLock(g_languageMutex);
        if (!g_cachedLanguageCode.empty() && haveWriteTime &&
            CompareFileTime(&currentWriteTime, &g_cachedIniWriteTime) == 0) {
            return;
        }
    }

    std::string detectedLanguageCode = DetectCurrentLanguageCodeFromIniPath(iniPath);
    if (detectedLanguageCode.empty()) {
        detectedLanguageCode = "eng";
    }

    std::lock_guard<std::mutex> cacheLock(g_languageMutex);
    g_defaultIniPath = iniPath;
    g_cachedLanguageCode = detectedLanguageCode;
    if (haveWriteTime) {
        g_cachedIniWriteTime = currentWriteTime;
    }
    else {
        ZeroMemory(&g_cachedIniWriteTime, sizeof(g_cachedIniWriteTime));
    }

    DbgLog("Language cache refreshed: ini='%s' lang='%s'\n", iniPath.c_str(), g_cachedLanguageCode.c_str());
}

} // namespace

std::atomic<bool> g_cancelRequested{ false };

const FieldDescriptor* GetFieldDescriptor(int fieldIndex)
{
    if (fieldIndex < 0 || fieldIndex >= FIELD_COUNT) return nullptr;
    return &kFieldDescriptors[fieldIndex];
}

const char* GetMultipleChoiceText(int fieldIndex, int value)
{
    switch (fieldIndex) {
    case FIELD_AUTO_UPDATE_STYLES:
        return value == 0 ? "Yes" : "No";
    case FIELD_TRACK_CHANGES_ENABLED_DISABLED:
        return value == 0 ? "Enabled" : "Disabled";
    case FIELD_ANONYMISATION:
        switch (value) {
        case 1: return "Personal information";
        case 2: return "Date and time";
        case 3: return "Personal information, Date and time";
        default: return "No";
        }
    default:
        return "";
    }
}

bool EnsureCachedParts(const std::string& ansiPath)
{
    std::lock_guard<std::mutex> lock(g_cacheMutex);
    if (g_cachedParts.path == ansiPath) return true;

    g_cachedParts = CachedParts();

    ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", g_cachedParts.coreXml);
    ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", g_cachedParts.appXml);
    ExtractFileFromZip(ansiPath.c_str(), "word/comments.xml", g_cachedParts.commentsXml);
    ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", g_cachedParts.settingsXml);
    ExtractFileFromZip(ansiPath.c_str(), "word/document.xml", g_cachedParts.documentXml);
    PopulateDerivedCachedParts(g_cachedParts);

    g_cachedParts.path = ansiPath;
    return true;
}

void ClearCache()
{
    std::lock_guard<std::mutex> lock(g_cacheMutex);
    g_cachedParts = CachedParts();
}

bool IsCanceled()
{
    return g_cancelRequested.load(std::memory_order_relaxed);
}

void SetPluginDefaultIniPath(const std::string& iniPath)
{
    std::string resolvedIniPath = ResolveWinCmdIniPath(iniPath);
    if (resolvedIniPath.empty()) {
        resolvedIniPath = GetDefaultIniPathFallback();
    }

    std::lock_guard<std::mutex> lock(g_languageMutex);
    g_defaultIniPath = resolvedIniPath;
    ZeroMemory(&g_cachedIniWriteTime, sizeof(g_cachedIniWriteTime));
    DbgLog("SetPluginDefaultIniPath: raw='%s' stored='%s'\n", iniPath.c_str(), g_defaultIniPath.c_str());
}

void RefreshPluginLanguageCache()
{
    RefreshPluginLanguageCacheIfNeeded();
}

std::string TranslatePluginText(const std::string& englishText)
{
    if (englishText.empty()) return englishText;

    std::wstring lngPath = GetPluginLngPath();
    if (lngPath.empty()) return englishText;

    RefreshPluginLanguageCacheIfNeeded();

    std::string langCode = "eng";
    {
        std::lock_guard<std::mutex> cacheLock(g_languageMutex);
        if (!g_cachedLanguageCode.empty()) {
            langCode = g_cachedLanguageCode;
        }
    }

    std::wstring key;
    if (!Utf8ToWideString(englishText, key)) return englishText;
    const wchar_t* missingSentinel = L"__PLUGIN_LNG_MISSING__";

    for (const std::string& candidate : GetLanguageSectionCandidates(langCode)) {
        std::wstring section(candidate.begin(), candidate.end());
        wchar_t translated[1024] = {};
        GetPrivateProfileStringW(section.c_str(), key.c_str(), missingSentinel, translated, static_cast<DWORD>(std::size(translated)), lngPath.c_str());
        std::wstring translatedValue(translated);
        if (!translatedValue.empty() && translatedValue != missingSentinel) {
            std::string translatedUtf8 = WideToUtf8(translatedValue);
            DbgLog("TranslatePluginText: lang='%s' candidate='%s' key='%s' -> '%s'\n",
                langCode.c_str(), candidate.c_str(), englishText.c_str(), translatedUtf8.c_str());
            return translatedUtf8;
        }
    }

    DbgLog("TranslatePluginText: lang='%s' key='%s' -> fallback\n", langCode.c_str(), englishText.c_str());
    return englishText;
}

std::string FileTimeToIso8601UTC(const FILETIME* ftUtc)
{
    if (!ftUtc) return std::string();
    SYSTEMTIME stUtc;
    if (!FileTimeToSystemTime(ftUtc, &stUtc)) return std::string();

    char buf[64];
    snprintf(buf, sizeof(buf), "%04d-%02d-%02dT%02d:%02d:%02dZ",
        stUtc.wYear, stUtc.wMonth, stUtc.wDay, stUtc.wHour, stUtc.wMinute, stUtc.wSecond);
    return std::string(buf);
}

bool ExtractFileFromZip(const char* zipPath, const char* fileNameInZip, std::string& output)
{
    mz_zip_archive zipArchive{};
    if (!mz_zip_reader_init_file(&zipArchive, zipPath, 0)) return false;

    int fileIndex = mz_zip_reader_locate_file(&zipArchive, fileNameInZip, nullptr, 0);
    if (fileIndex < 0) {
        mz_zip_reader_end(&zipArchive);
        return false;
    }

    size_t uncompressedSize = 0;
    void* p = mz_zip_reader_extract_file_to_heap(&zipArchive, fileNameInZip, &uncompressedSize, 0);
    if (!p) {
        mz_zip_reader_end(&zipArchive);
        return false;
    }

    output.assign(static_cast<char*>(p), uncompressedSize);
    mz_free(p);
    mz_zip_reader_end(&zipArchive);

    if (IsProbablyXml(fileNameInZip)) {
        if (!IsValidUtf8(output)) {
            std::string converted = AnsiToUtf8(output);
            if (!converted.empty()) output.swap(converted);
        }

        if (output.size() >= 3 &&
            static_cast<unsigned char>(output[0]) == 0xEF &&
            static_cast<unsigned char>(output[1]) == 0xBB &&
            static_cast<unsigned char>(output[2]) == 0xBF) {
            output.erase(0, 3);
        }

        if (output.rfind("<?xml", 0) != 0) {
            output = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + output;
        }
        else {
            size_t pos = output.find("?>");
            if (pos != std::string::npos) {
                std::string decl = output.substr(0, pos + 2);
                if (decl.find("UTF-8") == std::string::npos) {
                    output = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + output.substr(pos + 2);
                }
            }
        }
    }

    return true;
}

bool WidePathToAnsi(const WCHAR* wpath, std::string& out)
{
    if (!wpath) return false;
    int len = WideCharToMultiByte(CP_ACP, 0, wpath, -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return false;

    std::string buffer(static_cast<size_t>(len), '\0');
    if (WideCharToMultiByte(CP_ACP, 0, wpath, -1, &buffer[0], len, nullptr, nullptr) == 0) return false;

    out.assign(buffer.c_str());
    return true;
}

bool Utf8ToWideString(const std::string& in, std::wstring& out)
{
    if (in.empty()) {
        out.clear();
        return true;
    }

    int needed = MultiByteToWideChar(CP_UTF8, 0, in.c_str(), -1, nullptr, 0);
    if (needed <= 0) return false;

    std::wstring buffer(static_cast<size_t>(needed), L'\0');
    if (MultiByteToWideChar(CP_UTF8, 0, in.c_str(), -1, &buffer[0], needed) == 0) return false;

    out.assign(buffer.c_str());
    return true;
}

std::string WideToUtf8(const std::wstring& wstr)
{
    if (wstr.empty()) return std::string();
    int needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (needed <= 0) return std::string();

    std::string buffer(static_cast<size_t>(needed), '\0');
    if (WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &buffer[0], needed, nullptr, nullptr) == 0) return std::string();

    return std::string(buffer.c_str());
}

bool FormatSystemTimeToAnsi(const FILETIME& ftUtc, int unitIndex, char* outputStr, int maxLen)
{
    if (!outputStr || maxLen <= 0) return false;

    int maxWLen = maxLen;
    std::vector<wchar_t> wbuf(static_cast<size_t>(maxWLen), L'\0');
    if (!FormatSystemTimeToString(ftUtc, unitIndex, wbuf.data(), maxWLen)) return false;

    int res = WideCharToMultiByte(CP_ACP, 0, wbuf.data(), -1, outputStr, maxLen, nullptr, nullptr);
    if (res == 0) return false;
    outputStr[maxLen - 1] = '\0';
    return true;
}

std::set<std::string> GetTrackedChangeAuthorsFromAllXml(const char* zipPath)
{
    std::set<std::string> authors;
    mz_zip_archive zipArchive{};
    if (!mz_zip_reader_init_file(&zipArchive, zipPath, 0)) return authors;

    mz_uint numFiles = mz_zip_reader_get_num_files(&zipArchive);
    for (mz_uint i = 0; i < numFiles; ++i) {
        if (IsCanceled()) break;

        mz_zip_archive_file_stat fileStat;
        if (!mz_zip_reader_file_stat(&zipArchive, i, &fileStat)) continue;

        const char* fname = fileStat.m_filename;
        if (!fname) continue;
        if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr) continue;

        size_t uncompressedSize = 0;
        void* p = mz_zip_reader_extract_to_heap(&zipArchive, i, &uncompressedSize, 0);
        if (!p) continue;

        std::string content(static_cast<char*>(p), uncompressedSize);
        mz_free(p);

        tinyxml2::XMLDocument doc;
        if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS) continue;

        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) continue;

        ExtractAuthorsRecursive(root, authors);
    }

    mz_zip_reader_end(&zipArchive);
    return authors;
}

bool HasTrackedChanges(const char* zipPath)
{
    mz_zip_archive zipArchive{};
    if (!mz_zip_reader_init_file(&zipArchive, zipPath, 0)) return false;

    bool found = false;
    mz_uint numFiles = mz_zip_reader_get_num_files(&zipArchive);
    for (mz_uint i = 0; i < numFiles && !found; ++i) {
        if (IsCanceled()) break;

        mz_zip_archive_file_stat fileStat;
        if (!mz_zip_reader_file_stat(&zipArchive, i, &fileStat)) continue;

        const char* fname = fileStat.m_filename;
        if (!fname) continue;
        if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr) continue;

        size_t uncompressedSize = 0;
        void* p = mz_zip_reader_extract_to_heap(&zipArchive, i, &uncompressedSize, 0);
        if (!p) continue;

        std::string content(static_cast<char*>(p), uncompressedSize);
        mz_free(p);

        tinyxml2::XMLDocument doc;
        if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS) continue;

        std::function<void(tinyxml2::XMLElement*)> check = [&](tinyxml2::XMLElement* elem) {
            if (!elem || found) return;

            const char* name = elem->Name();
            if (name) {
                if (strcmp(name, "w:ins") == 0 ||
                    strcmp(name, "w:del") == 0 ||
                    strcmp(name, "w:moveFrom") == 0 ||
                    strcmp(name, "w:rPrChange") == 0 ||
                    strcmp(name, "w:pPrChange") == 0 ||
                    strcmp(name, "w:sectPrChange") == 0 ||
                    strcmp(name, "w:tblPrChange") == 0 ||
                    strcmp(name, "w:tblGridChange") == 0 ||
                    strcmp(name, "w:trPrChange") == 0 ||
                    strcmp(name, "w:tcPrChange") == 0) {
                    found = true;
                    return;
                }
            }

            for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child; child = child->NextSiblingElement()) {
                check(child);
            }
        };

        check(doc.RootElement());
    }

    mz_zip_reader_end(&zipArchive);
    return found;
}

int CountComments(const std::string& xmlContent)
{
    if (xmlContent.empty()) return 0;

    tinyxml2::XMLDocument doc;
    if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS) return 0;

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return 0;

    int count = 0;
    for (tinyxml2::XMLElement* comment = root->FirstChildElement("w:comment"); comment; comment = comment->NextSiblingElement("w:comment")) {
        ++count;
    }
    return count;
}

void PopulateDerivedCachedParts(CachedParts& parts)
{
    parts.coreTitle = GetXmlStringValue(parts.coreXml, "dc:title");
    parts.coreSubject = GetXmlStringValue(parts.coreXml, "dc:subject");
    parts.coreCreator = GetXmlStringValue(parts.coreXml, "dc:creator");
    parts.appManager = GetXmlStringValue(parts.appXml, "Manager");
    parts.appCompany = GetXmlStringValue(parts.appXml, "Company");
    parts.coreKeywords = GetXmlStringValue(parts.coreXml, "cp:keywords");
    parts.coreDescription = GetXmlStringValue(parts.coreXml, "dc:description");
    parts.appHyperlinkBase = GetXmlStringValue(parts.appXml, "HyperlinkBase");
    parts.appTemplate = GetXmlStringValue(parts.appXml, "Template");
    parts.coreCreated = GetXmlStringValue(parts.coreXml, "dcterms:created");
    parts.coreModified = GetXmlStringValue(parts.coreXml, "dcterms:modified");
    parts.coreLastPrinted = GetXmlStringValue(parts.coreXml, "cp:lastPrinted");
    parts.coreLastModifiedBy = GetXmlStringValue(parts.coreXml, "cp:lastModifiedBy");
    parts.coreRevision = GetXmlIntValue(parts.coreXml, "cp:revision");
    parts.appEditingTime = GetXmlIntValue(parts.appXml, "TotalTime");
    parts.appPages = GetXmlIntValue(parts.appXml, "Pages");
    parts.appParagraphs = GetXmlIntValue(parts.appXml, "Paragraphs");
    parts.appLines = GetXmlIntValue(parts.appXml, "Lines");
    parts.appWords = GetXmlIntValue(parts.appXml, "Words");
    parts.appCharacters = GetXmlIntValue(parts.appXml, "Characters");
    parts.compatibilityMode = IsCompatibilityModeEnabled(parts.settingsXml);
    parts.autoUpdateStyles = IsAutoUpdateStylesEnabled(parts.settingsXml);
    parts.anonymisationFlags = GetAnonymisedFlags(parts.settingsXml);
    parts.trackChangesEnabled = IsTrackChangesEnabled(parts.settingsXml);
    parts.commentsCount = CountComments(parts.commentsXml);
    parts.hiddenText = HasHiddenTextInDocumentXmlContent(parts.documentXml);
    parts.derivedReady = true;
}

bool IsTrackChangesEnabled(const std::string& settingsXmlContent)
{
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) return false;

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return false;

    return IsOnOffElementEnabled(root->FirstChildElement("w:trackRevisions"));
}

bool IsAutoUpdateStylesEnabled(const std::string& settingsXmlContent)
{
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) return false;

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return false;

    return IsOnOffElementEnabled(root->FirstChildElement("w:linkStyles"));
}

int GetAnonymisedFlags(const std::string& settingsXmlContent)
{
    if (settingsXmlContent.empty()) return 0;

    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) return 0;

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return 0;

    int flags = 0;
    if (IsOnOffElementEnabled(root->FirstChildElement("w:removePersonalInformation"))) flags |= 1;
    if (IsOnOffElementEnabled(root->FirstChildElement("w:removeDateAndTime"))) flags |= 2;
    return flags;
}

bool HasHiddenTextInDocumentXml(const char* zipPath)
{
    std::string xmlContent;
    if (!ExtractFileFromZip(zipPath, "word/document.xml", xmlContent)) return false;
    return HasHiddenTextInDocumentXmlContent(xmlContent);
}

bool HasHiddenTextInDocumentXmlContent(const std::string& xmlContent)
{
    if (xmlContent.empty()) return false;
    tinyxml2::XMLDocument doc;
    if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS) return false;

    tinyxml2::XMLElement* root = doc.FirstChildElement("w:document");
    if (!root) return false;

    tinyxml2::XMLElement* body = root->FirstChildElement("w:body");
    if (!body) return false;

    for (tinyxml2::XMLElement* para = body->FirstChildElement("w:p"); para; para = para->NextSiblingElement("w:p")) {
        for (tinyxml2::XMLElement* run = para->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r")) {
            tinyxml2::XMLElement* rPr = run->FirstChildElement("w:rPr");
            if (rPr && rPr->FirstChildElement("w:vanish")) return true;
        }
    }

    return false;
}

bool IsCompatibilityModeEnabled(const std::string& settingsXmlContent)
{
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) return false;

    tinyxml2::XMLElement* settings = doc.FirstChildElement("w:settings");
    if (!settings) return false;

    tinyxml2::XMLElement* compat = settings->FirstChildElement("w:compat");
    if (!compat) return false;

    for (tinyxml2::XMLElement* compatSetting = compat->FirstChildElement("w:compatSetting");
         compatSetting;
         compatSetting = compatSetting->NextSiblingElement("w:compatSetting")) {
        const char* nameAttr = compatSetting->Attribute("w:name");
        if (!nameAttr || strcmp(nameAttr, "compatibilityMode") != 0) continue;

        const char* valAttr = compatSetting->Attribute("w:val");
        if (!valAttr) return false;

        try {
            return std::stoi(valAttr) < 15;
        }
        catch (...) {
            return false;
        }
    }

    return false;
}

namespace {

std::string StripUtf8Bom(const std::string& xmlContent)
{
    if (xmlContent.size() >= 3 &&
        static_cast<unsigned char>(xmlContent[0]) == 0xEF &&
        static_cast<unsigned char>(xmlContent[1]) == 0xBB &&
        static_cast<unsigned char>(xmlContent[2]) == 0xBF) {
        return xmlContent.substr(3);
    }
    return xmlContent;
}

tinyxml2::XMLElement* FindElementByNameRecursive(tinyxml2::XMLElement* element, const char* elementName)
{
    if (!element || !elementName) return nullptr;

    if (const char* currentName = element->Name()) {
        if (strcmp(currentName, elementName) == 0) {
            return element;
        }
    }

    for (tinyxml2::XMLElement* child = element->FirstChildElement(); child; child = child->NextSiblingElement()) {
        if (tinyxml2::XMLElement* match = FindElementByNameRecursive(child, elementName)) {
            return match;
        }
    }

    return nullptr;
}

tinyxml2::XMLElement* FindXmlElement(tinyxml2::XMLDocument& doc, const char* elementName)
{
    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return nullptr;

    if (tinyxml2::XMLElement* directChild = root->FirstChildElement(elementName)) {
        return directChild;
    }

    return FindElementByNameRecursive(root, elementName);
}

bool ParseXmlDocument(const std::string& xmlContent, tinyxml2::XMLDocument& doc)
{
    if (xmlContent.empty()) return false;

    std::string normalized = StripUtf8Bom(xmlContent);
    return doc.Parse(normalized.c_str(), normalized.size()) == tinyxml2::XML_SUCCESS;
}

} // namespace

std::string GetXmlStringValue(const std::string& xmlContent, const char* elementName)
{
    tinyxml2::XMLDocument doc;
    if (!ParseXmlDocument(xmlContent, doc)) return std::string();

    tinyxml2::XMLElement* element = FindXmlElement(doc, elementName);
    if (element && element->GetText()) return element->GetText();
    return std::string();
}

int GetXmlIntValue(const std::string& xmlContent, const char* elementName)
{
    tinyxml2::XMLDocument doc;
    if (!ParseXmlDocument(xmlContent, doc)) return 0;

    tinyxml2::XMLElement* element = FindXmlElement(doc, elementName);
    if (!element) return 0;

    int value = 0;
    if (element->QueryIntText(&value) == tinyxml2::XML_SUCCESS) return value;
    return 0;
}

bool ParseIso8601ToFileTime(const std::string& iso8601Str, FILETIME* ftOut)
{
    if (iso8601Str.empty() || !ftOut) return false;

    SYSTEMTIME stUtc{};
    int year = 0;
    int month = 0;
    int day = 0;
    int hour = 0;
    int minute = 0;
    int second = 0;
    char tzBuffer[10]{};

    int resultCount = sscanf_s(iso8601Str.c_str(), "%d-%d-%dT%d:%d:%dZ", &year, &month, &day, &hour, &minute, &second);
    if (resultCount != 6) {
        resultCount = sscanf_s(iso8601Str.c_str(), "%d-%d-%dT%d:%d:%d%s", &year, &month, &day, &hour, &minute, &second, tzBuffer, static_cast<unsigned int>(sizeof(tzBuffer)));
        if (resultCount == 7) {
            int offsetH = 0;
            int offsetM = 0;
            char sign = tzBuffer[0];
            if (sscanf_s(tzBuffer + 1, "%d:%d", &offsetH, &offsetM) == 2) {
                if (sign == '+') {
                    hour -= offsetH;
                    minute -= offsetM;
                }
                else if (sign == '-') {
                    hour += offsetH;
                    minute += offsetM;
                }
            }
        }
        else {
            resultCount = sscanf_s(iso8601Str.c_str(), "%d-%d-%dT%d:%d:%d", &year, &month, &day, &hour, &minute, &second);
            if (resultCount != 6) {
                resultCount = sscanf_s(iso8601Str.c_str(), "%d-%d-%d", &year, &month, &day);
                if (resultCount != 3) return false;
                hour = 0;
                minute = 0;
                second = 0;
            }
        }
    }

    stUtc.wYear = static_cast<WORD>(year);
    stUtc.wMonth = static_cast<WORD>(month);
    stUtc.wDay = static_cast<WORD>(day);
    stUtc.wHour = static_cast<WORD>(hour);
    stUtc.wMinute = static_cast<WORD>(minute);
    stUtc.wSecond = static_cast<WORD>(second);

    return SystemTimeToFileTime(&stUtc, ftOut) != FALSE;
}

bool FormatSystemTimeToString(const FILETIME& ftUtc, int unitIndex, wchar_t* outputWStr, int maxWLen)
{
    if (!outputWStr || maxWLen <= 0) return false;
    outputWStr[0] = L'\0';

    FILETIME ftLocal;
    SYSTEMTIME sysTimeLocal;
    if (!FileTimeToLocalFileTime(&ftUtc, &ftLocal)) return false;
    if (!FileTimeToSystemTime(&ftLocal, &sysTimeLocal)) return false;

    int wlen = 0;
    switch (unitIndex) {
    case 0:
    case 3:
        wlen = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &sysTimeLocal, nullptr, outputWStr, maxWLen);
        if (wlen > 0) {
            int currentLen = static_cast<int>(wcslen(outputWStr));
            if (maxWLen - currentLen < 2) return false;
            wcscat_s(outputWStr, static_cast<size_t>(maxWLen), L" ");
            int timeLen = GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &sysTimeLocal, nullptr, outputWStr + currentLen + 1, maxWLen - (currentLen + 1));
            if (timeLen > 0) wlen += timeLen - 1;
            else return false;
        }
        break;
    case 1:
        wlen = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &sysTimeLocal, nullptr, outputWStr, maxWLen);
        break;
    case 2:
        wlen = GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &sysTimeLocal, nullptr, outputWStr, maxWLen);
        break;
    case 4:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%04d", sysTimeLocal.wYear);
        break;
    case 5:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTimeLocal.wMonth);
        break;
    case 6:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTimeLocal.wDay);
        break;
    case 7:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTimeLocal.wHour);
        break;
    case 8:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTimeLocal.wMinute);
        break;
    case 9:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTimeLocal.wSecond);
        break;
    case 10:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%04d-%02d-%02d", sysTimeLocal.wYear, sysTimeLocal.wMonth, sysTimeLocal.wDay);
        break;
    case 11:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d:%02d:%02d", sysTimeLocal.wHour, sysTimeLocal.wMinute, sysTimeLocal.wSecond);
        break;
    case 12:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%04d-%02d-%02d %02d:%02d:%02d",
            sysTimeLocal.wYear, sysTimeLocal.wMonth, sysTimeLocal.wDay,
            sysTimeLocal.wHour, sysTimeLocal.wMinute, sysTimeLocal.wSecond);
        break;
    default:
        return false;
    }

    if (wlen <= 0 || wlen >= maxWLen) return false;
    outputWStr[wlen] = L'\0';
    return true;
}

TrackedChangeCounts GetTrackedChangeCounts(const char* zipPath)
{
    TrackedChangeCounts counts;
    mz_zip_archive zipArchive{};
    if (!mz_zip_reader_init_file(&zipArchive, zipPath, 0)) return counts;

    mz_uint numFiles = mz_zip_reader_get_num_files(&zipArchive);
    for (mz_uint i = 0; i < numFiles; ++i) {
        if (IsCanceled()) break;

        mz_zip_archive_file_stat fileStat;
        if (!mz_zip_reader_file_stat(&zipArchive, i, &fileStat)) continue;

        const char* fname = fileStat.m_filename;
        if (!fname) continue;
        if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr) continue;

        size_t uncompressedSize = 0;
        void* p = mz_zip_reader_extract_to_heap(&zipArchive, i, &uncompressedSize, 0);
        if (!p) continue;

        std::string content(static_cast<char*>(p), uncompressedSize);
        mz_free(p);

        tinyxml2::XMLDocument doc;
        if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS) continue;

        CountTrackedChangesRecursive(doc.RootElement(), counts);
    }

    mz_zip_reader_end(&zipArchive);
    counts.totalRevisions = counts.insertions + counts.deletions + counts.moves + counts.formattingChanges;
    return counts;
}

FieldResult GetFieldResult(const std::string& ansiPath, int fieldIndex, int /*unitIndex*/)
{
    FieldResult res;

    size_t len = ansiPath.length();
    if (len < 5 || _stricmp(ansiPath.c_str() + len - 5, ".docx") != 0) return res;

    EnsureCachedParts(ansiPath);

    CachedParts cachedParts;

    {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        cachedParts = g_cachedParts;
    }

    TrackedChangeCounts trackedCounts;
    if (fieldIndex >= FIELD_TOTAL_REVISIONS && fieldIndex <= FIELD_TOTAL_FORMATTING_CHANGES) {
        if (IsCanceled()) return res;
        trackedCounts = GetTrackedChangeCounts(ansiPath.c_str());
    }

    switch (fieldIndex) {
    case FIELD_CORE_TITLE:
        res.s = cachedParts.coreTitle;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_CORE_SUBJECT:
        res.s = cachedParts.coreSubject;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_CORE_CREATOR:
        res.s = cachedParts.coreCreator;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_APP_MANAGER:
        res.s = cachedParts.appManager;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_APP_COMPANY:
        res.s = cachedParts.appCompany;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_CORE_KEYWORDS:
        res.s = cachedParts.coreKeywords;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_CORE_DESCRIPTION:
        res.s = cachedParts.coreDescription;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_APP_HYPERLINK_BASE:
        res.s = cachedParts.appHyperlinkBase;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_APP_TEMPLATE:
        res.s = cachedParts.appTemplate;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_CORE_CREATED_DATE:
    case FIELD_CORE_MODIFIED_DATE:
    case FIELD_CORE_LAST_PRINTED_DATE:
    {
        std::string dateStr;
        if (fieldIndex == FIELD_CORE_CREATED_DATE) dateStr = cachedParts.coreCreated;
        else if (fieldIndex == FIELD_CORE_MODIFIED_DATE) dateStr = cachedParts.coreModified;
        else dateStr = cachedParts.coreLastPrinted;

        if (dateStr.empty()) break;

        FILETIME ftUtc{};
        if (!ParseIso8601ToFileTime(dateStr, &ftUtc)) break;

        res.ft = ftUtc;
        res.type = FieldResult::FileTime;
        break;
    }
    case FIELD_CORE_LAST_MODIFIED_BY:
        res.s = cachedParts.coreLastModifiedBy;
        res.type = res.s.empty() ? FieldResult::Empty : FieldResult::StringUtf8;
        break;
    case FIELD_CORE_REVISION_NUMBER:
        res.i32 = cachedParts.coreRevision;
        res.type = FieldResult::Int32;
        break;
    case FIELD_APP_EDITING_TIME:
        res.i32 = cachedParts.appEditingTime;
        res.type = FieldResult::Int32;
        break;
    case FIELD_APP_PAGES:
        res.i32 = cachedParts.appPages;
        res.type = res.i32 == 0 ? FieldResult::Empty : FieldResult::Int32;
        break;
    case FIELD_APP_PARAGRAPHS:
        res.i32 = cachedParts.appParagraphs;
        res.type = res.i32 == 0 ? FieldResult::Empty : FieldResult::Int32;
        break;
    case FIELD_APP_LINES:
        res.i32 = cachedParts.appLines;
        res.type = res.i32 == 0 ? FieldResult::Empty : FieldResult::Int32;
        break;
    case FIELD_APP_WORDS:
        res.i32 = cachedParts.appWords;
        res.type = res.i32 == 0 ? FieldResult::Empty : FieldResult::Int32;
        break;
    case FIELD_APP_CHARACTERS:
        res.i32 = cachedParts.appCharacters;
        res.type = res.i32 == 0 ? FieldResult::Empty : FieldResult::Int32;
        break;
    case FIELD_COMPATMODE:
        res.b = cachedParts.compatibilityMode;
        res.type = FieldResult::Boolean;
        break;
    case FIELD_HIDDEN_TEXT:
        res.b = cachedParts.hiddenText;
        res.type = FieldResult::Boolean;
        break;
    case FIELD_COMMENTS:
        res.i32 = cachedParts.commentsCount;
        res.type = FieldResult::Int32;
        break;
    case FIELD_DOCUMENT_PROTECTION:
    {
        if (cachedParts.settingsXml.empty()) {
            res.s = TranslatePluginText("No protection");
            res.type = FieldResult::StringUtf8;
            break;
        }

        tinyxml2::XMLDocument doc;
        if (doc.Parse(cachedParts.settingsXml.c_str()) != tinyxml2::XML_SUCCESS) {
            res.s = TranslatePluginText("Error parsing settings.xml");
            res.type = FieldResult::StringUtf8;
            break;
        }

        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) {
            res.s = TranslatePluginText("No protection");
            res.type = FieldResult::StringUtf8;
            break;
        }

        tinyxml2::XMLElement* protectionElem = root->FirstChildElement("w:documentProtection");
        std::string protectionType = "No protection";
        if (protectionElem) {
            const char* enforcement = protectionElem->Attribute("w:enforcement");
            if (enforcement && strcmp(enforcement, "1") == 0) {
                const char* edit = protectionElem->Attribute("w:edit");
                if (edit) {
                    if (strcmp(edit, "readOnly") == 0) protectionType = "Read-only";
                    else if (strcmp(edit, "forms") == 0) protectionType = "Filling in forms";
                    else if (strcmp(edit, "comments") == 0) protectionType = "Comments";
                    else if (strcmp(edit, "trackedChanges") == 0) protectionType = "Tracked changes";
                    else protectionType = "Unknown protection type";
                }
                else {
                    protectionType = "Unknown protection type";
                }
            }
        }

        res.s = TranslatePluginText(protectionType);
        res.type = FieldResult::StringUtf8;
        break;
    }
    case FIELD_AUTO_UPDATE_STYLES:
        res.i32 = cachedParts.autoUpdateStyles ? 0 : 1;
        res.type = FieldResult::Int32;
        break;
    case FIELD_ANONYMISATION:
        res.i32 = cachedParts.anonymisationFlags;
        res.type = FieldResult::Int32;
        break;
    case FIELD_TRACKED_CHANGES:
        res.b = HasTrackedChanges(ansiPath.c_str());
        res.type = FieldResult::Boolean;
        break;
    case FIELD_TRACK_CHANGES_ENABLED_DISABLED:
        res.i32 = cachedParts.trackChangesEnabled ? 0 : 1;
        res.type = FieldResult::Int32;
        break;
    case FIELD_AUTHORS:
    {
        std::set<std::string> authors = GetTrackedChangeAuthorsFromAllXml(ansiPath.c_str());
        if (authors.empty()) break;

        std::stringstream ss;
        bool first = true;
        for (const std::string& author : authors) {
            if (!first) ss << ", ";
            ss << author;
            first = false;
        }

        res.s = ss.str();
        res.type = FieldResult::StringUtf8;
        break;
    }
    case FIELD_TOTAL_REVISIONS:
        res.i32 = trackedCounts.totalRevisions;
        res.type = FieldResult::Int32;
        break;
    case FIELD_TOTAL_INSERTIONS:
        res.i32 = trackedCounts.insertions;
        res.type = FieldResult::Int32;
        break;
    case FIELD_TOTAL_DELETIONS:
        res.i32 = trackedCounts.deletions;
        res.type = FieldResult::Int32;
        break;
    case FIELD_TOTAL_MOVES:
        res.i32 = trackedCounts.moves;
        res.type = FieldResult::Int32;
        break;
    case FIELD_TOTAL_FORMATTING_CHANGES:
        res.i32 = trackedCounts.formattingChanges;
        res.type = FieldResult::Int32;
        break;
    default:
        break;
    }

    return res;
}

void DbgLog(const char* fmt, ...)
{
    char tmpPath[MAX_PATH];
    if (!GetTempPathA(MAX_PATH, tmpPath)) return;

    std::string logPath = std::string(tmpPath) + "wdx_debug.log";
    FILE* f = nullptr;
    fopen_s(&f, logPath.c_str(), "a");
    if (!f) return;

    va_list args;
    va_start(args, fmt);
    vfprintf(f, fmt, args);
    va_end(args);
    fclose(f);
}

const char* GetIndirectAnsiChoiceText(const void* fieldValue)
{
    if (!fieldValue || reinterpret_cast<uintptr_t>(fieldValue) < 0xFFFF) return nullptr;

    const char* direct = static_cast<const char*>(fieldValue);
    if (direct[0] != '\0') {
        if (strcmp(direct, "Yes") == 0 || strcmp(direct, "No") == 0 ||
            strcmp(direct, "Enabled") == 0 || strcmp(direct, "Disabled") == 0 ||
            strcmp(direct, "Personal information") == 0 || strcmp(direct, "Date and time") == 0 ||
            strcmp(direct, "Personal information, Date and time") == 0) {
            return direct;
        }
    }

    const char* const* indirect = static_cast<const char* const*>(fieldValue);
    if (!indirect) return nullptr;

    const char* nested = *indirect;
    if (!nested || reinterpret_cast<uintptr_t>(nested) < 0xFFFF) return nullptr;
    return nested;
}

int NormalizeChoiceIndex(int fieldIndex, const void* fieldValue, int choiceCount)
{
    const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
    if (choiceText) {
        int mapped = MapChoiceTextToIndex(fieldIndex, choiceText);
        if (mapped >= 0) return mapped;
    }

    int rawIndex = GetChoiceIndex(fieldValue);
    if (rawIndex < 0 || choiceCount <= 0) return -1;
    if (rawIndex >= 0 && rawIndex < choiceCount) return rawIndex;
    if (rawIndex >= 1 && rawIndex <= choiceCount) return rawIndex - 1;
    return -1;
}

int GetContentValueWInternal(const std::string& ansiPath, int fieldIndex, int unitIndex, void* fieldValue, int maxLen)
{
    FieldResult r = GetFieldResult(ansiPath, fieldIndex, unitIndex);
    switch (r.type) {
    case FieldResult::Empty:
        return ft_fieldempty;
    case FieldResult::FileTime:
        if (unitIndex == 0) {
            memcpy(fieldValue, &r.ft, sizeof(FILETIME));
            return ft_datetime;
        }
        if (FormatSystemTimeToString(r.ft, unitIndex, static_cast<wchar_t*>(fieldValue), maxLen / static_cast<int>(sizeof(wchar_t)))) {
            return ft_stringw;
        }
        return ft_fieldempty;
    case FieldResult::Int32:
        if (fieldIndex == FIELD_AUTO_UPDATE_STYLES || fieldIndex == FIELD_ANONYMISATION || fieldIndex == FIELD_TRACK_CHANGES_ENABLED_DISABLED) {
            const char* choiceText = GetMultipleChoiceText(fieldIndex, r.i32);
            strncpy_s(static_cast<char*>(fieldValue), static_cast<size_t>(maxLen), choiceText, _TRUNCATE);
            return ft_multiplechoice;
        }
        *static_cast<int*>(fieldValue) = r.i32;
        return ft_numeric_32;
    case FieldResult::Int64:
        *static_cast<long long*>(fieldValue) = r.i64;
        return ft_numeric_64;
    case FieldResult::Boolean:
        *static_cast<int*>(fieldValue) = r.b ? 1 : 0;
        return ft_boolean;
    case FieldResult::StringUtf8:
    {
        if (r.s.empty() || !fieldValue || maxLen <= 0) return ft_fieldempty;
        std::wstring w;
        if (!Utf8ToWideString(r.s, w)) return ft_fieldempty;
        wcsncpy_s(static_cast<wchar_t*>(fieldValue), static_cast<size_t>(maxLen / sizeof(wchar_t)), w.c_str(), _TRUNCATE);
        return ft_stringw;
    }
    }

    return ft_fieldempty;
}

int GetContentValueInternal(const std::string& ansiPath, int fieldIndex, int unitIndex, void* fieldValue, int maxLen)
{
    FieldResult r = GetFieldResult(ansiPath, fieldIndex, unitIndex);
    switch (r.type) {
    case FieldResult::Empty:
        return ft_fieldempty;
    case FieldResult::FileTime:
        if (unitIndex == 0) {
            memcpy(fieldValue, &r.ft, sizeof(FILETIME));
            return ft_datetime;
        }
        if (FormatSystemTimeToAnsi(r.ft, unitIndex, static_cast<char*>(fieldValue), maxLen)) return ft_string;
        return ft_fieldempty;
    case FieldResult::Int32:
        if (fieldIndex == FIELD_AUTO_UPDATE_STYLES || fieldIndex == FIELD_ANONYMISATION || fieldIndex == FIELD_TRACK_CHANGES_ENABLED_DISABLED) {
            const char* choiceText = GetMultipleChoiceText(fieldIndex, r.i32);
            strncpy_s(static_cast<char*>(fieldValue), static_cast<size_t>(maxLen), choiceText, _TRUNCATE);
            return ft_multiplechoice;
        }
        *static_cast<int*>(fieldValue) = r.i32;
        return ft_numeric_32;
    case FieldResult::Int64:
        *static_cast<long long*>(fieldValue) = r.i64;
        return ft_numeric_64;
    case FieldResult::Boolean:
        *static_cast<int*>(fieldValue) = r.b ? 1 : 0;
        return ft_boolean;
    case FieldResult::StringUtf8:
    {
        if (r.s.empty() || !fieldValue || maxLen <= 0) return ft_fieldempty;
        std::wstring wide;
        if (!Utf8ToWideString(r.s, wide)) return ft_fieldempty;
        int res = WideCharToMultiByte(CP_ACP, 0, wide.c_str(), -1, static_cast<char*>(fieldValue), maxLen, nullptr, nullptr);
        if (res == 0) return ft_fieldempty;
        static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
        return ft_string;
    }
    }

    return ft_fieldempty;
}
