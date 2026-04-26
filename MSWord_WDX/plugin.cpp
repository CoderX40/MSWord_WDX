#include <windows.h>
#include <string>
#include <cstring>
#include <mutex>
#include <atomic>
#include <set>
#include <sstream>
#include <functional>
#include <cstddef>
#include <cstdlib>
#include <iomanip>
#include <ctime>
#include <cstdio>
#include <cstdarg>
#include <cctype>
#include <vector>
#include <map>
#include "miniz.h"
#include "tinyxml2.h"

#ifndef _WIN64
#pragma comment(linker, "/export:ContentGetSupportedField=_ContentGetSupportedField@16")
#pragma comment(linker, "/export:ContentGetValue=_ContentGetValue@24")
#pragma comment(linker, "/export:ContentGetValueW=_ContentGetValueW@24")
#pragma comment(linker, "/export:ContentGetSupportedFieldFlags=_ContentGetSupportedFieldFlags@4")
#pragma comment(linker, "/export:ContentSetValue=_ContentSetValue@24")
#pragma comment(linker, "/export:ContentSetValueW=_ContentSetValueW@24")
#pragma comment(linker, "/export:ContentEditValue=_ContentEditValue@32")
#pragma comment(linker, "/export:ContentEditValueW=_ContentEditValueW@32")
#endif

// Constants for Total Commander field types
#define ft_nomorefields     0
#define ft_numeric_32       1
#define ft_numeric_64       2
#define ft_numeric_floating 3
#define ft_date             4
#define ft_time             5
#define ft_boolean          6
#define ft_multiplechoice   7
#define ft_string           8
#define ft_fulltext         9
#define ft_datetime         10
#define ft_stringw          11
#define ft_fulltextw        12
#define ft_fieldempty       -3
#define ft_fileerror        -2
#define ft_nosuchfield      -1
#define ft_ondemand         -4
#define ft_notsupported     -5
#define ft_setsuccess       0
#define ft_setcancel        -2
#define ft_seterror         -1

// Constants for ContentGetSupportedFieldFlags
#define contflags_edit 1
#define contflags_substsize 2
#define contflags_substdatetime 4
#define contflags_substdate 6
#define contflags_substtime 8
#define contflags_substattributes 10
#define contflags_substattributestr 12
#define contflags_passthrough_size_float 14
#define contflags_substmask 14
#define contflags_fieldedit 16

// Field indices for plugin
enum {
    // Document Properties (Core & App)
    FIELD_CORE_TITLE = 0,
    FIELD_CORE_SUBJECT,
    FIELD_CORE_CREATOR,
    FIELD_APP_MANAGER,
    FIELD_APP_COMPANY,
    FIELD_CORE_KEYWORDS,
    FIELD_CORE_DESCRIPTION,
    FIELD_APP_HYPERLINK_BASE,
    FIELD_APP_TEMPLATE,
    // Statistics Fields
    FIELD_CORE_CREATED_DATE,
    FIELD_CORE_MODIFIED_DATE,
    FIELD_CORE_LAST_PRINTED_DATE,
    FIELD_CORE_LAST_MODIFIED_BY,
    FIELD_CORE_REVISION_NUMBER,
    FIELD_APP_EDITING_TIME,
    FIELD_APP_PAGES,
    FIELD_APP_PARAGRAPHS,
    FIELD_APP_LINES,
    FIELD_APP_WORDS,
    FIELD_APP_CHARACTERS,
    // Other Check Fields
    FIELD_COMPATMODE,
    FIELD_HIDDEN_TEXT,
    FIELD_COMMENTS,
    FIELD_DOCUMENT_PROTECTION,
    FIELD_AUTO_UPDATE_STYLES,
    FIELD_ANONYMISED_FILES,
    FIELD_TRACKED_CHANGES,
    FIELD_TCS_ON_OFF,
    FIELD_AUTHORS,
    FIELD_TOTAL_REVISIONS,
    FIELD_TOTAL_INSERTIONS,
    FIELD_TOTAL_DELETIONS,
    FIELD_TOTAL_MOVES,
    FIELD_TOTAL_FORMATTING_CHANGES,
    FIELD_COUNT
};

// --- Runtime globals for optional plugin features ---
static std::mutex g_cacheMutex;
struct CachedParts {
    std::string path;
    std::string coreXml;
    std::string appXml;
    std::string commentsXml;
    std::string settingsXml;
    std::string documentXml;
};
static CachedParts g_cachedParts;
static std::atomic<bool> g_cancelRequested{ false };


bool ExtractFileFromZip(const char* zipPath, const char* fileNameInZip, std::string& output);

bool EnsureCachedParts(const std::string& ansiPath)
{
    std::lock_guard<std::mutex> lock(g_cacheMutex);
    if (g_cachedParts.path == ansiPath)
        return true;

    g_cachedParts.path.clear();
    g_cachedParts.coreXml.clear();
    g_cachedParts.appXml.clear();
    g_cachedParts.commentsXml.clear();
    g_cachedParts.settingsXml.clear();
    g_cachedParts.documentXml.clear();

    ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", g_cachedParts.coreXml);
    ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", g_cachedParts.appXml);
    ExtractFileFromZip(ansiPath.c_str(), "word/comments.xml", g_cachedParts.commentsXml);
    ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", g_cachedParts.settingsXml);
    ExtractFileFromZip(ansiPath.c_str(), "word/document.xml", g_cachedParts.documentXml);

    g_cachedParts.path = ansiPath;
    return true;
}

static std::string FileTimeToIso8601UTC(const FILETIME* ftUtc)
{
    if (!ftUtc) return std::string();
    SYSTEMTIME stUtc;
    if (!FileTimeToSystemTime(ftUtc, &stUtc)) return std::string();
    char buf[64];
    snprintf(buf, sizeof(buf), "%04d-%02d-%02dT%02d:%02d:%02dZ",
        stUtc.wYear, stUtc.wMonth, stUtc.wDay, stUtc.wHour, stUtc.wMinute, stUtc.wSecond);
    return std::string(buf);
}

void ClearCache()
{
    std::lock_guard<std::mutex> lock(g_cacheMutex);
    g_cachedParts = CachedParts();
}

inline bool IsCanceled()
{
    return g_cancelRequested.load(std::memory_order_relaxed);
}

// --- ZIP extraction function using miniz ---
bool ExtractFileFromZip(const char* zipPath, const char* fileNameInZip, std::string& output)
{
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));

    if (!mz_zip_reader_init_file(&zip_archive, zipPath, 0))
        return false;

    int fileIndex = mz_zip_reader_locate_file(&zip_archive, fileNameInZip, nullptr, 0);
    if (fileIndex < 0)
    {
        mz_zip_reader_end(&zip_archive);
        return false;
    }
    size_t uncompressed_size = 0;
    void* p = mz_zip_reader_extract_file_to_heap(&zip_archive, fileNameInZip, &uncompressed_size, 0);
    if (!p)
    {
        mz_zip_reader_end(&zip_archive);
        return false;
    }

    output.assign(static_cast<char*>(p), uncompressed_size);
    mz_free(p);
    mz_zip_reader_end(&zip_archive);
    auto IsProbablyXml = [&](const char* name) {
        if (!name) return false;
        if (strstr(name, "docProps/") || strstr(name, "word/") || strcmp(name, "[Content_Types].xml") == 0) return true;
        size_t n = strlen(name);
        if (n > 4 && strcmp(name + n - 4, ".xml") == 0) return true;
        return false;
        };

    auto IsValidUtf8 = [](const std::string& str) -> bool {
        const unsigned char* bytes = (const unsigned char*)str.c_str();
        size_t len = str.size();
        size_t i = 0;
        while (i < len) {
            unsigned char c = bytes[i];
            if (c < 0x80) { i++; continue; }
            else if ((c >> 5) == 0x6) {
                if (i + 1 >= len) return false;
                if ((bytes[i + 1] >> 6) != 0x2) return false;
                i += 2;
            }
            else if ((c >> 4) == 0xE) {
                if (i + 2 >= len) return false;
                if ((bytes[i + 1] >> 6) != 0x2 || (bytes[i + 2] >> 6) != 0x2) return false;
                i += 3;
            }
            else if ((c >> 3) == 0x1E) {
                if (i + 3 >= len) return false;
                if ((bytes[i + 1] >> 6) != 0x2 || (bytes[i + 2] >> 6) != 0x2 || (bytes[i + 3] >> 6) != 0x2) return false;
                i += 4;
            }
            else return false;
        }
        return true;
        };

    auto AnsiToUtf8 = [](const std::string& in) -> std::string {
        if (in.empty()) return std::string();
        int wlen = MultiByteToWideChar(CP_ACP, 0, in.c_str(), (int)in.size(), NULL, 0);
        if (wlen == 0) return std::string();
        std::wstring w;
        w.resize(wlen);
        if (MultiByteToWideChar(CP_ACP, 0, in.c_str(), (int)in.size(), &w[0], wlen) == 0) return std::string();
        int ulen = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), wlen, NULL, 0, NULL, NULL);
        if (ulen == 0) return std::string();
        std::string out;
        out.resize(ulen);
        if (WideCharToMultiByte(CP_UTF8, 0, w.c_str(), wlen, &out[0], ulen, NULL, NULL) == 0) return std::string();
        return out;
        };

    if (IsProbablyXml(fileNameInZip)) {
        if (!IsValidUtf8(output)) {
            std::string converted = AnsiToUtf8(output);
            if (!converted.empty()) output.swap(converted);
        }
        if (output.rfind("<?xml", 0) != 0) {
            std::string decl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
            output = decl + output;
        }
        else {
            size_t pos = output.find("?>");
            if (pos != std::string::npos) {
                std::string decl = output.substr(0, pos + 2);
                if (decl.find("UTF-8") == std::string::npos) {
                    std::string rest = output.substr(pos + 2);
                    std::string newdecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
                    output = newdecl + rest;
                }
            }
        }
    }

    return true;
}

// --- ANSI / Unicode Conversions ---
bool WidePathToAnsi(const WCHAR* wpath, std::string& out)
{
    if (!wpath) return false;
    int len = WideCharToMultiByte(CP_ACP, 0, wpath, -1, NULL, 0, NULL, NULL);
    if (len <= 0) return false;
    out.resize(len - 1);
    if (WideCharToMultiByte(CP_ACP, 0, wpath, -1, &out[0], len, NULL, NULL) == 0) return false;
    return true;
}

bool Utf8ToWideString(const std::string& in, std::wstring& out)
{
    if (in.empty()) { out.clear(); return true; }
    int needed = MultiByteToWideChar(CP_UTF8, 0, in.c_str(), -1, NULL, 0);
    if (needed <= 0) return false;
    out.resize(needed - 1);
    if (MultiByteToWideChar(CP_UTF8, 0, in.c_str(), -1, &out[0], needed) == 0) return false;
    return true;
}

std::string WideToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (needed <= 0) return std::string();
    std::string out;
    out.resize(needed - 1);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &out[0], needed, NULL, NULL);
    return out;
}

bool FormatSystemTimeToString(const FILETIME& ft_utc, int unitIndex, wchar_t* outputWStr, int maxWLen);

bool FormatSystemTimeToAnsi(const FILETIME& ft_utc, int unitIndex, char* outputStr, int maxLen) {
    if (outputStr == nullptr || maxLen <= 0) return false;

    int maxWLen = maxLen / static_cast<int>(sizeof(wchar_t));
    if (maxWLen <= 0) return false;

    std::vector<wchar_t> wbuf(static_cast<size_t>(maxWLen));
    if (!FormatSystemTimeToString(ft_utc, unitIndex, wbuf.data(), maxWLen)) {
        return false;
    }

    int res = WideCharToMultiByte(CP_ACP, 0, wbuf.data(), -1, outputStr, maxLen, NULL, NULL);
    if (res == 0) return false;
    outputStr[maxLen - 1] = '\0';
    return true;
}

// --- XML parsing helpers using tinyxml2 ---
void ExtractAuthorsRecursive(tinyxml2::XMLElement* elem, std::set<std::string>& authors)
{
    if (!elem) return;

    const char* name = elem->Name();
    if (name)
    {
        // Add conditions for various formatting change tags
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
            strcmp(name, "w:tblLook") == 0)
        {
            const char* author = elem->Attribute("w:author");
            if (author)
                authors.insert(author);
            const char* originalAuthor = elem->Attribute("w:originalAuthor");
            if (originalAuthor)
                authors.insert(originalAuthor);
        }
    }

    for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child != nullptr; child = child->NextSiblingElement())
    {
        ExtractAuthorsRecursive(child, authors);
    }
}

// Retrieves all unique authors from tracked changes across all XML files in the "word/" directory.
std::set<std::string> GetTrackedChangeAuthorsFromAllXml(const char* zipPath)
{
    std::set<std::string> authors;

    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));

    if (!mz_zip_reader_init_file(&zip_archive, zipPath, 0))
        return authors;

    mz_uint num_files = mz_zip_reader_get_num_files(&zip_archive);
    for (mz_uint i = 0; i < num_files; ++i)
    {
        if (IsCanceled()) break;
        mz_zip_archive_file_stat file_stat;
        if (!mz_zip_reader_file_stat(&zip_archive, i, &file_stat))
            continue;

        const char* fname = file_stat.m_filename;
        if (!fname) continue;

        if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr)
            continue;

        size_t uncompressed_size = 0;
        void* p = mz_zip_reader_extract_to_heap(&zip_archive, i, &uncompressed_size, 0);
        if (!p) continue;

        std::string content(static_cast<char*>(p), uncompressed_size);
        mz_free(p);

        tinyxml2::XMLDocument doc;
        if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS) continue;

        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) continue;

        ExtractAuthorsRecursive(root, authors);
    }

    mz_zip_reader_end(&zip_archive);
    return authors;
}

// Checks if the document XML content contains any type of tracked changes (insertions, deletions, or formatting changes).
bool HasTrackedChanges(const char* zipPath)
{
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));

    if (!mz_zip_reader_init_file(&zip_archive, zipPath, 0))
        return false;

    bool found = false;
    mz_uint num_files = mz_zip_reader_get_num_files(&zip_archive);
    for (mz_uint i = 0; i < num_files; ++i)
    {
        if (IsCanceled()) break;
        mz_zip_archive_file_stat file_stat;
        if (!mz_zip_reader_file_stat(&zip_archive, i, &file_stat))
            continue;

        const char* fname = file_stat.m_filename;
        if (!fname) continue;

        // Process only XML files within the "word/" directory
        if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr)
            continue;

        size_t uncompressed_size = 0;
        void* p = mz_zip_reader_extract_to_heap(&zip_archive, i, &uncompressed_size, 0);
        if (!p) continue;

        std::string content(static_cast<char*>(p), uncompressed_size);
        mz_free(p);

        tinyxml2::XMLDocument doc;
        if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS) continue;

        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) continue;

        std::function<void(tinyxml2::XMLElement*)> check = [&](tinyxml2::XMLElement* elem)
            {
                if (!elem || found) return;

                const char* name = elem->Name();
                if (name)
                {
                    // ONLY check for explicit tracked change tags
                    if (strcmp(name, "w:ins") == 0 ||
                        strcmp(name, "w:del") == 0 ||
                        strcmp(name, "w:moveFrom") == 0 ||
                        strcmp(name, "w:rPrChange") == 0 ||     // Run properties (character formatting) change
                        strcmp(name, "w:pPrChange") == 0 ||     // Paragraph properties change
                        strcmp(name, "w:sectPrChange") == 0 ||  // Section properties change
                        strcmp(name, "w:tblPrChange") == 0 ||   // Table properties change
                        strcmp(name, "w:tblGridChange") == 0 || // Table grid properties change
                        strcmp(name, "w:trPrChange") == 0 ||    // Table row properties change
                        strcmp(name, "w:tcPrChange") == 0)      // Table cell properties change
                    {
                        found = true;
                        return;
                    }
                }

                for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child != nullptr; child = child->NextSiblingElement())
                    check(child);
            };

        check(root);
        if (found) break; // Found changes, no need to check further files
    }

    mz_zip_reader_end(&zip_archive);
    return found;
}

int CountComments(const std::string& xmlContent)
{
    if (xmlContent.empty()) return 0;

    tinyxml2::XMLDocument doc;
    if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS)
        return 0;

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return 0;

    int count = 0;
    for (tinyxml2::XMLElement* comment = root->FirstChildElement("w:comment"); comment != nullptr; comment = comment->NextSiblingElement("w:comment"))
    {
        count++;
    }
    return count;
}

static bool IsOnOffElementEnabled(tinyxml2::XMLElement* element)
{
    if (!element) return false;
    const char* val = element->Attribute("w:val");
    if (!val) return true;

    std::string s(val);
    for (char& ch : s)
        ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));

    return !(s == "0" || s == "false" || s == "off");
}

bool IsTrackChangesEnabled(const std::string& settingsXmlContent) {
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return false;
    }

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return false;

    return IsOnOffElementEnabled(root->FirstChildElement("w:trackRevisions"));
}

bool IsAutoUpdateStylesEnabled(const std::string& settingsXmlContent) {
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return false;
    }

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return false;

    return IsOnOffElementEnabled(root->FirstChildElement("w:linkStyles"));
}

int GetAnonymisedFlags(const std::string& settingsXmlContent) {
    if (settingsXmlContent.empty()) return 0;

    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return 0;
    }

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return 0;

    bool hasPI = IsOnOffElementEnabled(root->FirstChildElement("w:removePersonalInformation"));
    bool hasDate = IsOnOffElementEnabled(root->FirstChildElement("w:removeDateAndTime"));

    int flags = 0;
    if (hasPI) flags |= 1;
    if (hasDate) flags |= 2;
    return flags;
}

bool HasHiddenTextInDocumentXml(const char* zipPath)
{
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));

    if (!mz_zip_reader_init_file(&zip_archive, zipPath, 0))
        return false;

    const char* targetFile = "word/document.xml";
    int fileIndex = mz_zip_reader_locate_file(&zip_archive, targetFile, nullptr, 0);
    if (fileIndex < 0) {
        mz_zip_reader_end(&zip_archive);
        return false;
    }

    size_t uncompressed_size = 0;
    void* p = mz_zip_reader_extract_to_heap(&zip_archive, fileIndex, &uncompressed_size, 0);
    if (!p) {
        mz_zip_reader_end(&zip_archive);
        return false;
    }

    std::string xmlContent(static_cast<char*>(p), uncompressed_size);
    mz_free(p);
    mz_zip_reader_end(&zip_archive);

    tinyxml2::XMLDocument doc;
    tinyxml2::XMLError result = doc.Parse(xmlContent.c_str());
    if (result != tinyxml2::XML_SUCCESS)
        return false;

    tinyxml2::XMLElement* root = doc.FirstChildElement("w:document");
    if (!root) return false;

    tinyxml2::XMLElement* body = root->FirstChildElement("w:body");
    if (!body) return false;

    for (tinyxml2::XMLElement* para = body->FirstChildElement("w:p"); para; para = para->NextSiblingElement("w:p")) {
        for (tinyxml2::XMLElement* run = para->FirstChildElement("w:r"); run; run = run->NextSiblingElement("w:r")) {
            tinyxml2::XMLElement* rPr = run->FirstChildElement("w:rPr");
            if (!rPr) continue;

            if (rPr->FirstChildElement("w:vanish")) {
                return true;
            }
        }
    }

    return false;
}

bool IsCompatibilityModeEnabled(const std::string& settingsXmlContent)
{
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return false;
    }
    tinyxml2::XMLElement* settings = doc.FirstChildElement("w:settings");
    if (!settings) return false;

    tinyxml2::XMLElement* compat = settings->FirstChildElement("w:compat");
    if (!compat) {
        return false;
    }

    tinyxml2::XMLElement* compatSetting = compat->FirstChildElement("w:compatSetting");
    while (compatSetting) {
        const char* nameAttr = compatSetting->Attribute("w:name");
        if (nameAttr && strcmp(nameAttr, "compatibilityMode") == 0) {
            if (nameAttr && strcmp(nameAttr, "compatibilityMode") == 0) {
                const char* valAttr = compatSetting->Attribute("w:val");
                if (valAttr) {
                    try {
                        int compatVal = std::stoi(valAttr);
                        // Word 2013 (val="15") and newer are considered "non-compatibility mode"
                        // Word 2007 (val="12") and Word 2010 (val="14") are compatibility modes.
                        // Word 2003 (val="11") would also be compatibility mode.
                        return compatVal < 15;
                    }
                    catch (const std::invalid_argument&) {
                        return false;
                    }
                    catch (const std::out_of_range&) {
                        return false;
                    }
                }
            }
            compatSetting = compatSetting->NextSiblingElement("w:compatSetting");
        }

    }
    return false;
}

std::string GetXmlStringValue(const std::string& xmlContent, const char* elementName) {
    if (xmlContent.empty()) return "";
    tinyxml2::XMLDocument doc;
    if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS) return "";
    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return "";
    tinyxml2::XMLElement* element = root->FirstChildElement(elementName);
    if (element && element->GetText()) {
        return element->GetText();
    }
    return "";
}

int GetXmlIntValue(const std::string& xmlContent, const char* elementName) {
    if (xmlContent.empty()) return 0;
    tinyxml2::XMLDocument doc;
    if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS) return 0;
    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return 0;
    tinyxml2::XMLElement* element = root->FirstChildElement(elementName);
    if (element) {
        int val;
        if (element->QueryIntText(&val) == tinyxml2::XML_SUCCESS) {
            return val;
        }
    }
    return 0;
}

bool ParseIso8601ToFileTime(const std::string& iso8601_str, FILETIME* ft_out) {
    if (iso8601_str.empty() || ft_out == nullptr) {
        return false;
    }

    SYSTEMTIME st_utc = {};

    int year, month, day, hour = 0, minute = 0, second = 0;
    char tz_char_buffer[10];
    int result_count;

    result_count = sscanf_s(iso8601_str.c_str(), "%d-%d-%dT%d:%d:%dZ", &year, &month, &day, &hour, &minute, &second);
    if (result_count == 6) {
    }
    else {
        result_count = sscanf_s(iso8601_str.c_str(), "%d-%d-%dT%d:%d:%d%s", &year, &month, &day, &hour, &minute, &second, tz_char_buffer, (unsigned int)sizeof(tz_char_buffer));
        if (result_count == 7) {
            int offset_h = 0, offset_m = 0;
            char sign = tz_char_buffer[0];
            if (sscanf_s(tz_char_buffer + 1, "%d:%d", &offset_h, &offset_m) == 2) {
                if (sign == '+') {
                    hour -= offset_h;
                    minute -= offset_m;
                }
                else if (sign == '-') {
                    hour += offset_h;
                    minute += offset_m;
                }
            }
            else {
            }
        }
        else {
            result_count = sscanf_s(iso8601_str.c_str(), "%d-%d-%dT%d:%d:%d", &year, &month, &day, &hour, &minute, &second);
            if (result_count != 6) {
                result_count = sscanf_s(iso8601_str.c_str(), "%d-%d-%d", &year, &month, &day);
                if (result_count != 3) {
                    return false;
                }
                hour = 0; minute = 0; second = 0;
            }
        }
    }

    st_utc.wYear = year;
    st_utc.wMonth = month;
    st_utc.wDay = day;
    st_utc.wHour = hour;
    st_utc.wMinute = minute;
    st_utc.wSecond = second;

    if (!SystemTimeToFileTime(&st_utc, ft_out)) {
        return false;
    }

    return true;
}

bool FormatSystemTimeToString(const FILETIME& ft_utc, int unitIndex, wchar_t* outputWStr, int maxWLen) {
    outputWStr[0] = L'\0';

    SYSTEMTIME sysTime_local;

    FILETIME ft_local;
    if (!FileTimeToLocalFileTime(&ft_utc, &ft_local)) {
        return false;
    }
    if (!FileTimeToSystemTime(&ft_local, &sysTime_local)) {
        return false;
    }

    int wlen = 0;
    switch (unitIndex) {
    case 0:
    case 3:
        wlen = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &sysTime_local, NULL, outputWStr, maxWLen);
        if (wlen > 0) {
            // Explicitly cast to int as wcslen returns size_t
            int current_len = static_cast<int>(wcslen(outputWStr));
            if (maxWLen - current_len < 2) return false;
            wcscat_s(outputWStr, static_cast<size_t>(maxWLen), L" "); // wcscat_s expects size_t for buffer size

            // Explicitly cast to int for length calculation for GetTimeFormatW
            int time_len = GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &sysTime_local, NULL, outputWStr + current_len + 1, maxWLen - (current_len + 1));
            if (time_len > 0) wlen += (time_len - 1);
            else return false;
        }
        break;
    case 1:
        wlen = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &sysTime_local, NULL, outputWStr, maxWLen);
        break;
    case 2:
        wlen = GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &sysTime_local, NULL, outputWStr, maxWLen);
        break;
    case 4:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%04d", sysTime_local.wYear);
        break;
    case 5:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTime_local.wMonth);
        break;
    case 6:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTime_local.wDay);
        break;
    case 7:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTime_local.wHour);
        break;
    case 8:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTime_local.wMinute);
        break;
    case 9:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d", sysTime_local.wSecond);
        break;
    case 10:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%04d-%02d-%02d", sysTime_local.wYear, sysTime_local.wMonth, sysTime_local.wDay);
        break;
    case 11:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%02d:%02d:%02d", sysTime_local.wHour, sysTime_local.wMinute, sysTime_local.wSecond);
        break;
    case 12:
        wlen = swprintf_s(outputWStr, static_cast<size_t>(maxWLen), L"%04d-%02d-%02d %02d:%02d:%02d",
            sysTime_local.wYear, sysTime_local.wMonth, sysTime_local.wDay,
            sysTime_local.wHour, sysTime_local.wMinute, sysTime_local.wSecond);
        break;
    default:
        return false;
    }

    if (wlen <= 0 || wlen >= maxWLen) {
        return false;
    }
    outputWStr[wlen] = L'\0';
    return true;
}

bool ReplaceFileInZip(const char* zipPath, const char* entryName, const std::string& data)
{
    mz_zip_archive zip = {};

    if (!mz_zip_reader_init_file(&zip, zipPath, 0))
        return false;

    mz_zip_archive newZip = {};
    if (!mz_zip_writer_init_heap(&newZip, 0, 0))
        return false;

    int fileCount = (int)mz_zip_reader_get_num_files(&zip);

    for (int i = 0; i < fileCount; i++) {
        mz_zip_archive_file_stat fileStat;
        mz_zip_reader_file_stat(&zip, i, &fileStat);

        if (strcmp(fileStat.m_filename, entryName) == 0) {
            mz_zip_writer_add_mem(&newZip, entryName, data.c_str(), data.size(), MZ_DEFAULT_COMPRESSION);
        }
        else {
            size_t size = 0;
            void* buf = mz_zip_reader_extract_to_heap(&zip, i, &size, 0);
            mz_zip_writer_add_mem(&newZip, fileStat.m_filename, buf, size, MZ_DEFAULT_COMPRESSION);
            mz_free(buf);
        }
    }

    mz_zip_reader_end(&zip);

    size_t newSize = 0;
    void* newBuf = nullptr;

    if (!mz_zip_writer_finalize_heap_archive(&newZip, &newBuf, &newSize)) {
        mz_zip_writer_end(&newZip);
        return false;
    }

    mz_zip_writer_end(&newZip);

    FILE* f = nullptr;
    errno_t err = fopen_s(&f, zipPath, "wb");

    if (err != 0 || !f) {
        mz_free(newBuf);
        return false;
    }

    size_t written = fwrite(newBuf, 1, newSize, f);
    fclose(f);

    mz_free(newBuf);

    if (written != newSize) {
        return false;
    }

    return true;
}

struct TrackedChangeCounts {
    int insertions = 0;
    int deletions = 0;
    int moves = 0;
    int formattingChanges = 0;
    int totalRevisions = 0;
    std::set<std::string> uniqueFormattingChanges; // To count formatting changes like Word
};

void CountTrackedChangesRecursive(tinyxml2::XMLElement* elem, TrackedChangeCounts& counts) {
    if (!elem) return;

    if (IsCanceled()) return;

    const char* name = elem->Name();
    if (name) {
        if (strcmp(name, "w:ins") == 0) {
            counts.insertions++;
        }
        else if (strcmp(name, "w:del") == 0) {
            counts.deletions++;
        }
        else if (strcmp(name, "w:moveFrom") == 0) {
            counts.moves++;
        }
        else {
            // Strictly check for explicit formatting change tracking tags
            std::string change_id = name;
            bool isFormattingChange = false;

            if (strcmp(name, "w:rPrChange") == 0 ||     // Run properties (character formatting) change
                strcmp(name, "w:pPrChange") == 0 ||     // Paragraph properties change
                strcmp(name, "w:sectPrChange") == 0 ||  // Section properties change
                strcmp(name, "w:tblPrChange") == 0 ||   // Table properties change
                strcmp(name, "w:tblGridChange") == 0 || // Table grid properties change
                strcmp(name, "w:trPrChange") == 0 ||    // Table row properties change
                strcmp(name, "w:tcPrChange") == 0)      // Table cell properties change
            {
                isFormattingChange = true;
                if (elem->FirstChildElement()) {
                    change_id += ":" + std::string(elem->FirstChildElement()->Name());
                }
                else {
                    change_id += ":noChild";
                }
            }

            if (isFormattingChange) {
                if (counts.uniqueFormattingChanges.find(change_id) == counts.uniqueFormattingChanges.end()) {
                    counts.uniqueFormattingChanges.insert(change_id);
                    counts.formattingChanges++;
                }
            }
        }
    }

    for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child != nullptr; child = child->NextSiblingElement()) {
        CountTrackedChangesRecursive(child, counts);
    }
}

TrackedChangeCounts GetTrackedChangeCounts(const char* zipPath) {
    TrackedChangeCounts counts;
    mz_zip_archive zip_archive;
    memset(&zip_archive, 0, sizeof(zip_archive));

    if (!mz_zip_reader_init_file(&zip_archive, zipPath, 0))
        return counts;

    mz_uint num_files = mz_zip_reader_get_num_files(&zip_archive);
    for (mz_uint i = 0; i < num_files; ++i)
    {
        if (IsCanceled()) break;
        mz_zip_archive_file_stat file_stat;
        if (!mz_zip_reader_file_stat(&zip_archive, i, &file_stat))
            continue;

        const char* fname = file_stat.m_filename;
        if (!fname) continue;

        if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr)
            continue;

        size_t uncompressed_size = 0;
        void* p = mz_zip_reader_extract_to_heap(&zip_archive, i, &uncompressed_size, 0);
        if (!p) continue;

        std::string content(static_cast<char*>(p), uncompressed_size);
        mz_free(p);

        tinyxml2::XMLDocument doc;
        if (doc.Parse(content.c_str()) != tinyxml2::XML_SUCCESS) continue;

        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) continue;

        CountTrackedChangesRecursive(root, counts);
    }
    mz_zip_reader_end(&zip_archive);
    counts.totalRevisions = counts.insertions + counts.deletions + counts.moves + counts.formattingChanges;
    return counts;
}

struct FieldResult {
    enum Type { Empty, FileTime, Int32, Int64, Boolean, StringUtf8 } type = Empty;
    FILETIME ft{};
    int32_t i32 = 0;
    int64_t i64 = 0;
    bool b = false;
    std::string s;
};

FieldResult GetFieldResult(const std::string& ansiPath, int fieldIndex, int unitIndex)
{
    FieldResult res;

    size_t len = ansiPath.length();
    if (len < 5 || _stricmp(ansiPath.c_str() + len - 5, ".docx") != 0)
        return res;

    std::string documentXml;
    std::string commentsXml;
    std::string settingsXml;
    std::string coreXml;
    std::string appXml;

    EnsureCachedParts(ansiPath);

    {
        std::lock_guard<std::mutex> lock(g_cacheMutex);
        coreXml = g_cachedParts.coreXml;
        appXml = g_cachedParts.appXml;
        commentsXml = g_cachedParts.commentsXml;
        settingsXml = g_cachedParts.settingsXml;
        documentXml = g_cachedParts.documentXml;
    }

    bool needsCoreXml = false;
    bool needsAppXml = false;
    bool needsWordXml = false;
    bool needsCommentsXml = false;
    bool needsSettingsXml = false;

    switch (fieldIndex) {
    case FIELD_CORE_TITLE:
    case FIELD_CORE_SUBJECT:
    case FIELD_CORE_CREATOR:
    case FIELD_CORE_KEYWORDS:
    case FIELD_CORE_DESCRIPTION:
    case FIELD_CORE_LAST_MODIFIED_BY:
    case FIELD_CORE_CREATED_DATE:
    case FIELD_CORE_MODIFIED_DATE:
    case FIELD_CORE_LAST_PRINTED_DATE:
    case FIELD_CORE_REVISION_NUMBER:
        needsCoreXml = true; break;
    case FIELD_APP_MANAGER:
    case FIELD_APP_COMPANY:
    case FIELD_APP_HYPERLINK_BASE:
    case FIELD_APP_PAGES:
    case FIELD_APP_WORDS:
    case FIELD_APP_CHARACTERS:
    case FIELD_APP_LINES:
    case FIELD_APP_PARAGRAPHS:
    case FIELD_APP_EDITING_TIME:
        needsAppXml = true; break;
    case FIELD_COMPATMODE:
    case FIELD_COMMENTS:
        needsCommentsXml = true; break;
    case FIELD_DOCUMENT_PROTECTION:
    case FIELD_AUTO_UPDATE_STYLES:
    case FIELD_ANONYMISED_FILES:
    case FIELD_TCS_ON_OFF:
        needsSettingsXml = true; break;
    case FIELD_HIDDEN_TEXT:
    case FIELD_TRACKED_CHANGES:
    case FIELD_TOTAL_REVISIONS:
    case FIELD_TOTAL_INSERTIONS:
    case FIELD_TOTAL_DELETIONS:
    case FIELD_TOTAL_MOVES:
    case FIELD_TOTAL_FORMATTING_CHANGES:
        needsWordXml = true; break;
    case FIELD_AUTHORS:
    default:
        break;
    }

    if (needsCoreXml && coreXml.empty()) {
        if (!ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            res.type = FieldResult::Empty; return res;
        }
    }
    if (needsAppXml && appXml.empty()) {
        if (!ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
            res.type = FieldResult::Empty; return res;
        }
    }
    if (needsCommentsXml && commentsXml.empty()) {
        ExtractFileFromZip(ansiPath.c_str(), "word/comments.xml", commentsXml);
    }
    if (needsSettingsXml && settingsXml.empty()) {
        if (!ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml)) {
            res.type = FieldResult::Empty; return res;
        }
    }
    if (needsWordXml && documentXml.empty()) {
        if (!ExtractFileFromZip(ansiPath.c_str(), "word/document.xml", documentXml)) {
            res.type = FieldResult::Empty; return res;
        }
    }

    TrackedChangeCounts trackedCounts;
    if (fieldIndex >= FIELD_TOTAL_REVISIONS && fieldIndex <= FIELD_TOTAL_FORMATTING_CHANGES) {
        if (IsCanceled()) { res.type = FieldResult::Empty; return res; }
        trackedCounts = GetTrackedChangeCounts(ansiPath.c_str());
    }

    switch (fieldIndex) {
    case FIELD_CORE_TITLE:
        res.s = GetXmlStringValue(coreXml, "dc:title"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_CORE_SUBJECT:
        res.s = GetXmlStringValue(coreXml, "dc:subject"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_CORE_CREATOR:
        res.s = GetXmlStringValue(coreXml, "dc:creator"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_APP_MANAGER:
        res.s = GetXmlStringValue(appXml, "Manager"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_APP_COMPANY:
        res.s = GetXmlStringValue(appXml, "Company"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_CORE_KEYWORDS:
        res.s = GetXmlStringValue(coreXml, "cp:keywords"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_CORE_DESCRIPTION:
        res.s = GetXmlStringValue(coreXml, "dc:description"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_APP_HYPERLINK_BASE:
        res.s = GetXmlStringValue(appXml, "HyperlinkBase"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_APP_TEMPLATE:
        res.s = GetXmlStringValue(appXml, "Template"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_CORE_CREATED_DATE:
    case FIELD_CORE_MODIFIED_DATE:
    case FIELD_CORE_LAST_PRINTED_DATE:
    {
        std::string dateStr;
        if (fieldIndex == FIELD_CORE_CREATED_DATE) dateStr = GetXmlStringValue(coreXml, "dcterms:created");
        else if (fieldIndex == FIELD_CORE_MODIFIED_DATE) dateStr = GetXmlStringValue(coreXml, "dcterms:modified");
        else if (fieldIndex == FIELD_CORE_LAST_PRINTED_DATE) dateStr = GetXmlStringValue(coreXml, "cp:lastPrinted");
        if (dateStr.empty()) { res.type = FieldResult::Empty; break; }
        FILETIME ft_utc;
        if (!ParseIso8601ToFileTime(dateStr, &ft_utc)) { res.type = FieldResult::Empty; break; }
        res.ft = ft_utc; res.type = FieldResult::FileTime; break;
    }
    case FIELD_CORE_LAST_MODIFIED_BY:
        res.s = GetXmlStringValue(coreXml, "cp:lastModifiedBy"); if (res.s.empty()) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::StringUtf8; break;
    case FIELD_CORE_REVISION_NUMBER:
        res.i32 = GetXmlIntValue(coreXml, "cp:revision"); res.type = FieldResult::Int32; break;
    case FIELD_APP_EDITING_TIME:
        res.i32 = GetXmlIntValue(appXml, "TotalTime"); res.type = FieldResult::Int32; break;
    case FIELD_APP_PAGES:
        res.i32 = GetXmlIntValue(appXml, "Pages"); if (res.i32 == 0) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::Int32; break;
    case FIELD_APP_PARAGRAPHS:
        res.i32 = GetXmlIntValue(appXml, "Paragraphs"); if (res.i32 == 0) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::Int32; break;
    case FIELD_APP_LINES:
        res.i32 = GetXmlIntValue(appXml, "Lines"); if (res.i32 == 0) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::Int32; break;
    case FIELD_APP_WORDS:
        res.i32 = GetXmlIntValue(appXml, "Words"); if (res.i32 == 0) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::Int32; break;
    case FIELD_APP_CHARACTERS:
        res.i32 = GetXmlIntValue(appXml, "Characters"); if (res.i32 == 0) { res.type = FieldResult::Empty; }
        else res.type = FieldResult::Int32; break;
    case FIELD_COMPATMODE:
        res.b = IsCompatibilityModeEnabled(settingsXml); res.type = FieldResult::Boolean; break;
    case FIELD_HIDDEN_TEXT:
        res.b = HasHiddenTextInDocumentXml(ansiPath.c_str()); res.type = FieldResult::Boolean; break;
    case FIELD_COMMENTS:
        res.i32 = CountComments(commentsXml); res.type = FieldResult::Int32; break;
    case FIELD_DOCUMENT_PROTECTION:
    {
        if (settingsXml.empty()) { res.s = "No protection"; res.type = FieldResult::StringUtf8; break; }
        tinyxml2::XMLDocument doc;
        if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS) { res.s = "Error parsing settings.xml"; res.type = FieldResult::StringUtf8; break; }
        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) { res.s = "No protection"; res.type = FieldResult::StringUtf8; break; }
        tinyxml2::XMLElement* protectionElem = root->FirstChildElement("w:documentProtection");
        std::string protectionType = "No protection";
        if (protectionElem) {
            const char* enforcement = protectionElem->Attribute("w:enforcement");
            if (enforcement && strcmp(enforcement, "1") == 0) {
                const char* edit = protectionElem->Attribute("w:edit");
                if (edit) {
                    if (strcmp(edit, "readOnly") == 0) protectionType = "Read-Only";
                    else if (strcmp(edit, "forms") == 0) protectionType = "Forms";
                    else if (strcmp(edit, "comments") == 0) protectionType = "Comments";
                    else if (strcmp(edit, "trackedChanges") == 0) protectionType = "Tracked Changes";
                    else protectionType = "Unknown protection type";
                }
                else protectionType = "Unknown protection type";
            }
        }
        res.s = protectionType; res.type = FieldResult::StringUtf8; break;
    }
    case FIELD_AUTO_UPDATE_STYLES:
        res.i32 = IsAutoUpdateStylesEnabled(settingsXml) ? 0 : 1;
        res.type = FieldResult::Int32; break;
    case FIELD_ANONYMISED_FILES:
        res.i32 = GetAnonymisedFlags(settingsXml);
        res.type = FieldResult::Int32; break;
    case FIELD_TRACKED_CHANGES:
        res.b = HasTrackedChanges(ansiPath.c_str()); res.type = FieldResult::Boolean; break;
    case FIELD_TCS_ON_OFF:
        res.i32 = IsTrackChangesEnabled(settingsXml) ? 0 : 1;
        res.type = FieldResult::Int32; break;
    case FIELD_AUTHORS:
    {
        auto authors = GetTrackedChangeAuthorsFromAllXml(ansiPath.c_str());
        if (authors.empty()) { res.type = FieldResult::Empty; break; }
        std::stringstream ss;
        bool first = true;
        for (const auto& author : authors) { if (!first) ss << ", "; ss << author; first = false; }
        res.s = ss.str(); res.type = FieldResult::StringUtf8; break;
    }
    case FIELD_TOTAL_REVISIONS: res.i32 = trackedCounts.totalRevisions; res.type = FieldResult::Int32; break;
    case FIELD_TOTAL_INSERTIONS: res.i32 = trackedCounts.insertions; res.type = FieldResult::Int32; break;
    case FIELD_TOTAL_DELETIONS: res.i32 = trackedCounts.deletions; res.type = FieldResult::Int32; break;
    case FIELD_TOTAL_MOVES: res.i32 = trackedCounts.moves; res.type = FieldResult::Int32; break;
    case FIELD_TOTAL_FORMATTING_CHANGES: res.i32 = trackedCounts.formattingChanges; res.type = FieldResult::Int32; break;
    default: res.type = FieldResult::Empty; break;
    }

    return res;
}


// --- Total Commander Content Plugin API ---

// Debug logging — writes to %TEMP%\wdx_debug.log
static void DbgLog(const char* fmt, ...)
{
    char tmpPath[MAX_PATH];
    if (!GetTempPathA(MAX_PATH, tmpPath)) return;
    std::string logPath = std::string(tmpPath) + "wdx_debug.log";
    FILE* f = nullptr;
    fopen_s(&f, logPath.c_str(), "a");
    if (!f) return;
    va_list args; va_start(args, fmt);
    vfprintf(f, fmt, args);
    va_end(args);
    fclose(f);
}

static int GetChoiceIndex(const void* fieldValue)
{
    // TC may pass a small integer token instead of a pointer for ft_multiplechoice.
    if (fieldValue == nullptr || reinterpret_cast<uintptr_t>(fieldValue) < 0xFFFF) {
        if (fieldValue == nullptr) return -1;
        return static_cast<int>(reinterpret_cast<uintptr_t>(fieldValue));
    }

    return -1;
}

static int MapChoiceTextToIndex(int fieldIndex, const char* text)
{
    if (!text || !*text) return -1;

    if (fieldIndex == FIELD_AUTO_UPDATE_STYLES) {
        if (strcmp(text, "Yes") == 0) return 0;
        if (strcmp(text, "No") == 0) return 1;
    }
    else if (fieldIndex == FIELD_TCS_ON_OFF) {
        if (strcmp(text, "Activated") == 0) return 0;
        if (strcmp(text, "Deactivated") == 0) return 1;
    }
    else if (fieldIndex == FIELD_ANONYMISED_FILES) {
        if (strcmp(text, "No") == 0) return 0;
        if (strcmp(text, "Personal Information") == 0) return 1;
        if (strcmp(text, "Date and Time") == 0) return 2;
        if (strcmp(text, "Personal Information and Date and Time") == 0) return 3;
    }

    return -1;
}

static const char* GetIndirectAnsiChoiceText(const void* fieldValue)
{
    if (!fieldValue || reinterpret_cast<uintptr_t>(fieldValue) < 0xFFFF)
        return nullptr;

    const char* direct = static_cast<const char*>(fieldValue);
    if (direct[0] != '\0') {
        if (strcmp(direct, "Yes") == 0 || strcmp(direct, "No") == 0 ||
            strcmp(direct, "Activated") == 0 || strcmp(direct, "Deactivated") == 0 ||
            strcmp(direct, "Personal Information") == 0 || strcmp(direct, "Date and Time") == 0 ||
            strcmp(direct, "Personal Information and Date and Time") == 0) {
            return direct;
        }
    }

    const char* const* indirect = static_cast<const char* const*>(fieldValue);
    if (!indirect) return nullptr;

    const char* nested = *indirect;
    if (!nested || reinterpret_cast<uintptr_t>(nested) < 0xFFFF)
        return nullptr;

    return nested;
}

static int NormalizeChoiceIndex(int fieldIndex, const void* fieldValue, int choiceCount)
{
    const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
    if (choiceText) {
        int mapped = MapChoiceTextToIndex(fieldIndex, choiceText);
        if (mapped >= 0)
            return mapped;
    }

    int rawIndex = GetChoiceIndex(fieldValue);
    if (rawIndex < 0 || choiceCount <= 0)
        return -1;

    if (rawIndex >= 0 && rawIndex < choiceCount)
        return rawIndex;
    if (rawIndex >= 1 && rawIndex <= choiceCount)
        return rawIndex - 1;
    return -1;
}

extern "C" {

    __declspec(dllexport) int __stdcall ContentGetSupportedField(int fieldIndex, char* fieldName, char* units, int maxLen)
    {
        switch (fieldIndex)
        {
            // Document Properties
        case FIELD_CORE_TITLE:
            strncpy_s(fieldName, maxLen, "Document Title", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_CORE_SUBJECT:
            strncpy_s(fieldName, maxLen, "Subject", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_CORE_CREATOR:
            strncpy_s(fieldName, maxLen, "Author", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_APP_MANAGER:
            strncpy_s(fieldName, maxLen, "Manager", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_APP_COMPANY:
            strncpy_s(fieldName, maxLen, "Company", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_CORE_KEYWORDS:
            strncpy_s(fieldName, maxLen, "Keywords", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_CORE_DESCRIPTION:
            strncpy_s(fieldName, maxLen, "Comments", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_APP_HYPERLINK_BASE:
            strncpy_s(fieldName, maxLen, "Hyperlink base", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_APP_TEMPLATE:
            strncpy_s(fieldName, maxLen, "Template", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;

            // Statistics Fields
        case FIELD_CORE_CREATED_DATE:
            strncpy_s(fieldName, maxLen, "Created", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_datetime;
        case FIELD_CORE_MODIFIED_DATE:
            strncpy_s(fieldName, maxLen, "Modified", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_datetime;
        case FIELD_CORE_LAST_PRINTED_DATE:
            strncpy_s(fieldName, maxLen, "Printed", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_datetime;
        case FIELD_CORE_LAST_MODIFIED_BY:
            strncpy_s(fieldName, maxLen, "Last saved by", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_CORE_REVISION_NUMBER:
            strncpy_s(fieldName, maxLen, "Revision number", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_APP_EDITING_TIME:
            strncpy_s(fieldName, maxLen, "Total editing time", _TRUNCATE);
            strncpy_s(units, maxLen, "min", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_APP_PAGES:
            strncpy_s(fieldName, maxLen, "Pages", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_APP_PARAGRAPHS:
            strncpy_s(fieldName, maxLen, "Paragraphs", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_APP_LINES:
            strncpy_s(fieldName, maxLen, "Lines", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_APP_WORDS:
            strncpy_s(fieldName, maxLen, "Words", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_APP_CHARACTERS:
            strncpy_s(fieldName, maxLen, "Characters", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;

            // Other Check Fields
        case FIELD_COMPATMODE:
            strncpy_s(fieldName, maxLen, "Compatibility mode", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_boolean;
        case FIELD_HIDDEN_TEXT:
            strncpy_s(fieldName, maxLen, "Hidden text", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_boolean;
        case FIELD_COMMENTS:
            strncpy_s(fieldName, maxLen, "Number of comments", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_DOCUMENT_PROTECTION:
            strncpy_s(fieldName, maxLen, "Document Protection", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_AUTO_UPDATE_STYLES:
            strncpy_s(fieldName, maxLen, "Auto Update Styles", _TRUNCATE);
            strncpy_s(units, maxLen, "Yes|No", _TRUNCATE);
            return ft_multiplechoice;
        case FIELD_ANONYMISED_FILES:
            strncpy_s(fieldName, maxLen, "Files Anonymised", _TRUNCATE);
            strncpy_s(units, maxLen, "No|Personal Information|Date and Time|Personal Information and Date and Time", _TRUNCATE);
            return ft_multiplechoice;
        case FIELD_TCS_ON_OFF:
            strncpy_s(fieldName, maxLen, "Track Changes", _TRUNCATE);
            strncpy_s(units, maxLen, "Activated|Deactivated", _TRUNCATE);
            return ft_multiplechoice;
        case FIELD_TRACKED_CHANGES:
            strncpy_s(fieldName, maxLen, "Tracked Changes Present in Document", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_boolean;
        case FIELD_AUTHORS:
            strncpy_s(fieldName, maxLen, "Tracked Changes Authors", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_stringw;
        case FIELD_TOTAL_REVISIONS:
            strncpy_s(fieldName, maxLen, "Total Revisions", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_TOTAL_INSERTIONS:
            strncpy_s(fieldName, maxLen, "Total Insertions", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_TOTAL_DELETIONS:
            strncpy_s(fieldName, maxLen, "Total Deletions", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_TOTAL_MOVES:
            strncpy_s(fieldName, maxLen, "Total Moves", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;
        case FIELD_TOTAL_FORMATTING_CHANGES:
            strncpy_s(fieldName, maxLen, "Total Formatting Changes", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_numeric_32;

        default:
            return ft_nomorefields;
        }
    }


    static bool SaveReplacementsToDocx(const std::string& path, const std::map<std::string, std::string>& replacements)
    {
        std::string tmpPath = path + ".tmp";

        mz_zip_archive reader;
        memset(&reader, 0, sizeof(reader));
        if (!mz_zip_reader_init_file(&reader, path.c_str(), 0)) {
            return false;
        }

        mz_zip_archive writer;
        memset(&writer, 0, sizeof(writer));
        if (!mz_zip_writer_init_file(&writer, tmpPath.c_str(), 0)) {
            mz_zip_reader_end(&reader);
            return false;
        }

        mz_uint num_files = mz_zip_reader_get_num_files(&reader);
        for (mz_uint i = 0; i < num_files; ++i) {
            mz_zip_archive_file_stat stat;
            if (!mz_zip_reader_file_stat(&reader, i, &stat)) continue;
            const char* fname = stat.m_filename;
            if (!fname) continue;
            std::string name(fname);

            auto it = replacements.find(name);
            if (it != replacements.end()) {
                const std::string& data = it->second;
                if (!mz_zip_writer_add_mem(&writer, name.c_str(), data.data(), data.size(), MZ_DEFAULT_COMPRESSION)) {
                    mz_zip_end(&writer);
                    mz_zip_reader_end(&reader);
                    return false;
                }
            }
            else {
                if (!mz_zip_writer_add_from_zip_reader(&writer, &reader, i)) {
                    mz_zip_end(&writer);
                    mz_zip_reader_end(&reader);
                    return false;
                }
            }
        }

        for (const auto& kv : replacements) {
            bool exists = false;
            mz_uint count = mz_zip_reader_get_num_files(&reader);
            for (mz_uint i = 0; i < count; ++i) {
                mz_zip_archive_file_stat st;
                if (!mz_zip_reader_file_stat(&reader, i, &st)) continue;
                if (st.m_filename && kv.first == st.m_filename) { exists = true; break; }
            }
            if (!exists) {
                if (!mz_zip_writer_add_mem(&writer, kv.first.c_str(), kv.second.data(), kv.second.size(), MZ_DEFAULT_COMPRESSION)) {
                    mz_zip_end(&writer);
                    mz_zip_reader_end(&reader);
                    return false;
                }
            }
        }

        mz_zip_writer_finalize_archive(&writer);
        mz_zip_writer_end(&writer);
        mz_zip_reader_end(&reader);

        mz_zip_archive tmp_reader;
        memset(&tmp_reader, 0, sizeof(tmp_reader));
        bool tmp_ok = false;
        if (mz_zip_reader_init_file(&tmp_reader, tmpPath.c_str(), 0)) {
            mz_uint tmp_num = mz_zip_reader_get_num_files(&tmp_reader);
            if (tmp_num > 0) {
                int idx = mz_zip_reader_locate_file(&tmp_reader, "[Content_Types].xml", nullptr, 0);
                if (idx >= 0) tmp_ok = true;
            }
            mz_zip_reader_end(&tmp_reader);
        }


        if (!MoveFileExA(tmpPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED)) {
            return false;
        }
        return true;
    }


    __declspec(dllexport) int __stdcall ContentGetValueW(
        WCHAR* fileName, int fieldIndex, int unitIndex,
        void* fieldValue, int maxLen, int flags)
    {
        if (!fileName) return ft_fieldempty;
        std::string ansiPath;
        if (!WidePathToAnsi(fileName, ansiPath)) return ft_fieldempty;

        g_cancelRequested.store(false, std::memory_order_relaxed);

        auto WideToUtf8 = [](const std::wstring& w) -> std::string {
            if (w.empty()) return std::string();
            int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
            if (n <= 0) return std::string();
            std::string s(n - 1, '\0');
            WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], n, nullptr, nullptr);
            return s;
            };

        std::string coreXml, appXml, settingsXml, documentXml, commentsXml;
        ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml);
        ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml);
        ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml);
        ExtractFileFromZip(ansiPath.c_str(), "word/document.xml", documentXml);
        ExtractFileFromZip(ansiPath.c_str(), "word/comments.xml", commentsXml);

        FieldResult r = GetFieldResult(ansiPath, fieldIndex, unitIndex);
        switch (r.type) {
        case FieldResult::Empty: return ft_fieldempty;
        case FieldResult::FileTime:
            if (unitIndex == 0) { memcpy(fieldValue, &r.ft, sizeof(FILETIME)); return ft_datetime; }
            if (FormatSystemTimeToString(r.ft, unitIndex, static_cast<wchar_t*>(fieldValue), maxLen / sizeof(wchar_t))) return ft_stringw;
            return ft_fieldempty;
        case FieldResult::Int32:
            if (fieldIndex == FIELD_AUTO_UPDATE_STYLES ||
                fieldIndex == FIELD_ANONYMISED_FILES ||
                fieldIndex == FIELD_TCS_ON_OFF)
            {
                const char* choiceText = "";
                if (fieldIndex == FIELD_AUTO_UPDATE_STYLES)
                    choiceText = (r.i32 == 0) ? "Yes" : "No";
                else if (fieldIndex == FIELD_TCS_ON_OFF)
                    choiceText = (r.i32 == 0) ? "Activated" : "Deactivated";
                else { // FIELD_ANONYMISED_FILES
                    switch (r.i32) {
                    case 1: choiceText = "Personal Information"; break;
                    case 2: choiceText = "Date and Time"; break;
                    case 3: choiceText = "Personal Information and Date and Time"; break;
                    default: choiceText = "No"; break;
                    }
                }
                strncpy_s(static_cast<char*>(fieldValue), static_cast<size_t>(maxLen), choiceText, _TRUNCATE);
                return ft_multiplechoice;
            }
            *(int*)fieldValue = r.i32;
            return ft_numeric_32;
        case FieldResult::Int64: *(long long*)fieldValue = r.i64; return ft_numeric_64;
        case FieldResult::Boolean: *((int*)fieldValue) = r.b ? 1 : 0; return ft_boolean;
        case FieldResult::StringUtf8:
        {
            if (r.s.empty()) { return ft_fieldempty; }
            if (!fieldValue || maxLen <= 0) return ft_fieldempty;

            std::wstring w;
            Utf8ToWideString(r.s, w);
            wcsncpy_s(static_cast<wchar_t*>(fieldValue), static_cast<size_t>(maxLen / sizeof(wchar_t)), w.c_str(), _TRUNCATE);
            static_cast<wchar_t*>(fieldValue)[maxLen / sizeof(wchar_t) - 1] = L'\0';
            return ft_stringw;
        }
        }
        return ft_fieldempty;
    }

    __declspec(dllexport) int __stdcall ContentGetValue(
        char* fileName, int fieldIndex, int unitIndex,
        void* fieldValue, int maxLen, int flags)
    {
        if (!fileName) return ft_fieldempty;
        std::string path = fileName;

        g_cancelRequested.store(false, std::memory_order_relaxed);

        FieldResult r = GetFieldResult(path, fieldIndex, unitIndex);
        switch (r.type) {
        case FieldResult::Empty: return ft_fieldempty;
        case FieldResult::FileTime:
            if (unitIndex == 0) { memcpy(fieldValue, &r.ft, sizeof(FILETIME)); return ft_datetime; }
            if (FormatSystemTimeToAnsi(r.ft, unitIndex, static_cast<char*>(fieldValue), maxLen)) return ft_string;
            return ft_fieldempty;
        case FieldResult::Int32:
            if (fieldIndex == FIELD_AUTO_UPDATE_STYLES ||
                fieldIndex == FIELD_ANONYMISED_FILES ||
                fieldIndex == FIELD_TCS_ON_OFF)
            {
                const char* choiceText = "";
                if (fieldIndex == FIELD_AUTO_UPDATE_STYLES)
                    choiceText = (r.i32 == 0) ? "Yes" : "No";
                else if (fieldIndex == FIELD_TCS_ON_OFF)
                    choiceText = (r.i32 == 0) ? "Activated" : "Deactivated";
                else {
                    switch (r.i32) {
                    case 1: choiceText = "Personal Information"; break;
                    case 2: choiceText = "Date and Time"; break;
                    case 3: choiceText = "Personal Information and Date and Time"; break;
                    default: choiceText = "No"; break;
                    }
                }
                strncpy_s(static_cast<char*>(fieldValue), static_cast<size_t>(maxLen), choiceText, _TRUNCATE);
                return ft_multiplechoice;
            }
            *(int*)fieldValue = r.i32;
            return ft_numeric_32;
        case FieldResult::Int64: *(long long*)fieldValue = r.i64; return ft_numeric_64;
        case FieldResult::Boolean: *((int*)fieldValue) = r.b ? 1 : 0; return ft_boolean;
        case FieldResult::StringUtf8:
            if (r.s.empty()) { return ft_fieldempty; }
            if (!fieldValue || maxLen <= 0) return ft_fieldempty;
            {
                int wneeded = MultiByteToWideChar(CP_UTF8, 0, r.s.c_str(), -1, NULL, 0);
                if (wneeded <= 0) return ft_fieldempty;
                std::wstring w;
                w.resize(wneeded - 1);
                if (MultiByteToWideChar(CP_UTF8, 0, r.s.c_str(), -1, &w[0], wneeded) == 0) return ft_fieldempty;

                if (maxLen <= 0 || fieldValue == nullptr) return ft_fieldempty;
                int res = WideCharToMultiByte(CP_ACP, 0, w.c_str(), -1, static_cast<char*>(fieldValue), maxLen, NULL, NULL);
                if (res == 0) return ft_fieldempty;
                static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
                return ft_string;
            }
        }
        return ft_fieldempty;
    }

    __declspec(dllexport) int __stdcall ContentGetDetectString(char* DetectString, int maxlen)
    {
        if (!DetectString || maxlen <= 0) return 1;
        const char* shortOnly = "EXT=docx";
        size_t needShort = strlen(shortOnly) + 1;
        if (static_cast<size_t>(maxlen) < needShort) return 1;
        strcpy_s(DetectString, static_cast<size_t>(maxlen), shortOnly);
        return 0;
    }

    __declspec(dllexport) int __stdcall ContentGetDetectStringW(WCHAR* DetectString, int maxlen)
    {
        if (!DetectString || maxlen <= 0) return 1;
        const std::wstring shortOnly = L"EXT=docx";
        size_t needShort = shortOnly.length() + 1;
        if (static_cast<size_t>(maxlen) < needShort) return 1;
        wcscpy_s(DetectString, static_cast<size_t>(maxlen), shortOnly.c_str());
        return 0;
    }

    static wchar_t* SafeGetIniPath(void* dparm)
    {
        if (!dparm) return nullptr;

#if defined(_MSC_VER)
        __try {
            void* pRaw = nullptr;
            memcpy(&pRaw, reinterpret_cast<char*>(dparm) + sizeof(void*), sizeof(void*));
            if (pRaw == nullptr) return nullptr;

            wchar_t* pw = static_cast<wchar_t*>(pRaw);
            size_t maxCheck = 4096;
            size_t len = 0;
            __try {
                for (; len < maxCheck; ++len) {
                    if (pw[len] == L'\0') break;
                }
            }
            __except (EXCEPTION_EXECUTE_HANDLER) {
                return nullptr;
            }
            if (len == maxCheck) return nullptr;
            return pw;
        }
        __except (EXCEPTION_EXECUTE_HANDLER) {
            return nullptr;
        }
#else
        void* pRaw = nullptr;
        memcpy(&pRaw, reinterpret_cast<char*>(dparm) + sizeof(void*), sizeof(void*));
        if (!pRaw) return nullptr;
        wchar_t* pw = static_cast<wchar_t*>(pRaw);
        if (pw[0] == L'\0') return nullptr;
        return pw;
#endif
    }

    __declspec(dllexport) void __stdcall ContentSetDefaultParams(void* dparm)
    {
        if (!dparm) return;
        wchar_t* pIniPath = SafeGetIniPath(dparm);

        if (pIniPath) {
        }
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
        if (fieldIndex == -1) {
            return contflags_edit | contflags_fieldedit;
        }

        switch (fieldIndex) {
        case FIELD_CORE_TITLE:
        case FIELD_CORE_SUBJECT:
        case FIELD_CORE_CREATOR:
        case FIELD_APP_MANAGER:
        case FIELD_APP_COMPANY:
        case FIELD_CORE_KEYWORDS:
        case FIELD_CORE_DESCRIPTION:
        case FIELD_APP_HYPERLINK_BASE:
        case FIELD_CORE_CREATED_DATE:
        case FIELD_CORE_MODIFIED_DATE:
        case FIELD_CORE_LAST_PRINTED_DATE:
        case FIELD_CORE_LAST_MODIFIED_BY:
        case FIELD_CORE_REVISION_NUMBER:
        case FIELD_APP_EDITING_TIME:
        case FIELD_AUTO_UPDATE_STYLES:
        case FIELD_ANONYMISED_FILES:
        case FIELD_TCS_ON_OFF:
            return contflags_edit;
        case FIELD_DOCUMENT_PROTECTION:
        case FIELD_AUTHORS:
            return contflags_edit | contflags_fieldedit;
        default:
            return 0;
        }
    }

    __declspec(dllexport) void __stdcall ContentPluginUnloading(void)
    {
        g_cancelRequested.store(true, std::memory_order_relaxed);
        ClearCache();
    }

    // --- Setting Values for Core and App XML ---

    bool SetXmlStringValue(std::string& xmlContent, const char* elementName, const std::string& value) {
        if (xmlContent.empty() || !elementName || elementName[0] == '\0') return false;

        std::string xmlDecl;
        if (xmlContent.rfind("<?xml", 0) == 0) {
            size_t pos = xmlContent.find("?>");
            if (pos != std::string::npos) {
                xmlDecl = xmlContent.substr(0, pos + 2);
            }
        }

        tinyxml2::XMLDocument doc;
        if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
            return false;
        }

        tinyxml2::XMLElement* root = doc.RootElement();
        if (!root) return false;

        tinyxml2::XMLElement* element = root->FirstChildElement(elementName);
        if (element) {
            element->SetText(value.c_str());
        }
        else {
            const char* colon = strchr(elementName, ':');
            if (colon) {
                std::string prefix(elementName, colon - elementName);
                std::string xmlnsAttr = std::string("xmlns:") + prefix;
                const char* existing = root->Attribute(xmlnsAttr.c_str());
                if (!existing) {
                    if (prefix == "dc") root->SetAttribute(xmlnsAttr.c_str(), "http://purl.org/dc/elements/1.1/");
                    else if (prefix == "cp") root->SetAttribute(xmlnsAttr.c_str(), "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                    else if (prefix == "dcterms") root->SetAttribute(xmlnsAttr.c_str(), "http://purl.org/dc/terms/");
                    else if (prefix == "w") root->SetAttribute(xmlnsAttr.c_str(), "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                }
            }

            element = doc.NewElement(elementName);
            element->SetText(value.c_str());
            root->InsertEndChild(element);
        }

        tinyxml2::XMLPrinter printer(nullptr, true);
        doc.Accept(&printer);
        std::string printed = printer.CStr();

        if (!xmlDecl.empty() && printed.rfind("<?xml", 0) != 0) {
            xmlContent = xmlDecl + "\n" + printed;
        }
        else {
            xmlContent = printed;
        }

        return true;
    }

    bool SaveXmlToZip(const char* zipPath, const char* fileNameInZip, const std::string& content) {
        std::map<std::string, std::string> replacements;
        replacements[std::string(fileNameInZip)] = content;
        return SaveReplacementsToDocx(std::string(zipPath), replacements);
    }

    // ----------------------------------------------------------------
    // Author rename helper: walks all word/*.xml files and renames
    // w:author / w:originalAuthor attributes.
    // oldAuthor empty = rename all authors.
    // ----------------------------------------------------------------
    static bool RenameTrackedChangeAuthors(const std::string& ansiPath,
        const std::string& oldAuthor,
        const std::string& newAuthor)
    {
        mz_zip_archive reader;
        memset(&reader, 0, sizeof(reader));
        if (!mz_zip_reader_init_file(&reader, ansiPath.c_str(), 0)) return false;

        std::map<std::string, std::string> replacements;
        mz_uint num_files = mz_zip_reader_get_num_files(&reader);
        for (mz_uint i = 0; i < num_files; ++i) {
            mz_zip_archive_file_stat stat;
            if (!mz_zip_reader_file_stat(&reader, i, &stat)) continue;
            const char* fname = stat.m_filename;
            if (!fname) continue;
            if (strncmp(fname, "word/", 5) != 0 || strstr(fname, ".xml") == nullptr) continue;

            size_t sz = 0;
            void* p = mz_zip_reader_extract_to_heap(&reader, i, &sz, 0);
            if (!p) continue;
            std::string xml(static_cast<char*>(p), sz);
            mz_free(p);

            tinyxml2::XMLDocument doc;
            if (doc.Parse(xml.c_str()) != tinyxml2::XML_SUCCESS) continue;

            bool modified = false;
            std::function<void(tinyxml2::XMLElement*)> walk = [&](tinyxml2::XMLElement* elem) {
                if (!elem) return;
                for (const char* attr : { "w:author", "w:originalAuthor" }) {
                    const char* val = elem->Attribute(attr);
                    if (!val) continue;
                    if (oldAuthor.empty() || oldAuthor == val) {
                        elem->SetAttribute(attr, newAuthor.c_str());
                        modified = true;
                    }
                }
                for (auto* child = elem->FirstChildElement(); child; child = child->NextSiblingElement())
                    walk(child);
                };
            walk(doc.RootElement());

            if (modified) {
                tinyxml2::XMLPrinter printer(nullptr, true);
                doc.Accept(&printer);
                replacements[std::string(fname)] = printer.CStr();
            }
        }
        mz_zip_reader_end(&reader);
        if (replacements.empty()) return true;
        return SaveReplacementsToDocx(ansiPath, replacements);
    }

    // ----------------------------------------------------------------
    // Author rename dialog (built in memory, no .rc file needed)
    // ----------------------------------------------------------------
    struct AuthorDlgData {
        std::wstring oldAuthor;
        std::wstring newAuthor;
        bool accepted = false;
    };

#define ADL_OLD_EDIT 101
#define ADL_NEW_EDIT 102

    static INT_PTR CALLBACK AuthorDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
    {
        switch (msg) {
        case WM_INITDIALOG:
            SetWindowLongPtr(hDlg, GWLP_USERDATA, lParam);
            SetWindowTextW(hDlg, L"Rename Tracked Change Author");
            return TRUE;
        case WM_COMMAND:
            switch (LOWORD(wParam)) {
            case IDOK:
            {
                AuthorDlgData* d = reinterpret_cast<AuthorDlgData*>(GetWindowLongPtr(hDlg, GWLP_USERDATA));
                wchar_t oldName[512] = {};
                wchar_t newName[512] = {};
                GetWindowTextW(GetDlgItem(hDlg, ADL_OLD_EDIT), oldName, 512);
                GetWindowTextW(GetDlgItem(hDlg, ADL_NEW_EDIT), newName, 512);

                if (newName[0] == L'\0') {
                    MessageBoxW(hDlg, L"Please enter a new name.", L"Error", MB_OK);
                    return TRUE;
                }

                d->oldAuthor = oldName;
                d->newAuthor = newName;
                d->accepted = true;
                EndDialog(hDlg, IDOK);
                return TRUE;
            }
            case IDCANCEL:
                EndDialog(hDlg, IDCANCEL);
                return TRUE;
            }
            break;
        case WM_CLOSE:
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        return FALSE;
    }

    static DLGTEMPLATE* BuildAuthorDlgTemplate()
    {
        static WORD buf[2048];
        memset(buf, 0, sizeof(buf));
        WORD* p = buf;

        auto W = [&](WORD w) { *p++ = w; };
        auto DW = [&](DWORD dw) { memcpy(p, &dw, 4); p += 2; };
        auto WS = [&](const wchar_t* s) { while (*s) *p++ = (WORD)*s++; *p++ = 0; };
        auto Al = [&]() { if ((DWORD_PTR)p & 2) *p++ = 0; };

        // 6 controls, 260x120 DLUs
        DW(DS_SETFONT | DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU);
        DW(0); W(6); W(0); W(0); W(260); W(120);
        W(0); W(0); WS(L"Rename Author"); W(8); WS(L"MS Shell Dlg");

        auto Item = [&](DWORD style, short x, short y, short w, short h, WORD id,
            const wchar_t* cls, const wchar_t* text) {
                Al();
                DW(style | WS_CHILD | WS_VISIBLE); DW(0);
                W((WORD)x); W((WORD)y); W((WORD)w); W((WORD)h); W(id);
                WS(cls); WS(text); W(0);
            };

        Item(SS_LEFT, 7, 7, 246, 8, 0xFFFF, L"STATIC", L"Old Author Name (leave blank to rename ALL authors):");
        Item(ES_AUTOHSCROLL | WS_BORDER, 7, 18, 246, 14, ADL_OLD_EDIT, L"EDIT", L"");
        Item(SS_LEFT, 7, 40, 246, 8, 0xFFFF, L"STATIC", L"New Author Name:");
        Item(ES_AUTOHSCROLL | WS_BORDER, 7, 51, 246, 14, ADL_NEW_EDIT, L"EDIT", L"");
        Item(BS_DEFPUSHBUTTON, 148, 80, 50, 14, IDOK, L"BUTTON", L"OK");
        Item(BS_PUSHBUTTON, 203, 80, 50, 14, IDCANCEL, L"BUTTON", L"Cancel");

        return reinterpret_cast<DLGTEMPLATE*>(buf);
    }

    // ----------------------------------------------------------------
    // Document protection dialog
    // ----------------------------------------------------------------
    struct ProtDlgData {
        bool currentlyProtected = false; // true if file already has enforcement=1
        std::wstring chosenMode;         // "No protection" | "Read-Only" | etc.
        std::wstring password;
        bool accepted = false;
    };

#define PDL_COMBO   201
#define PDL_PASSLBL 202
#define PDL_PASS    203

    static INT_PTR CALLBACK ProtDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
    {
        switch (msg) {
        case WM_INITDIALOG:
        {
            SetWindowLongPtr(hDlg, GWLP_USERDATA, lParam);
            ProtDlgData* d = reinterpret_cast<ProtDlgData*>(lParam);
            SetWindowTextW(hDlg, L"Document Protection");
            HWND hCombo = GetDlgItem(hDlg, PDL_COMBO);
            // Always offer "No protection" (remove/clear enforcement)
            SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"No protection");
            if (!d->currentlyProtected) {
                // Only offer protective modes when file is currently unprotected —
                // we can't know the existing password so we can't modify protection in-place
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Read-Only");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Forms");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Comments");
                SendMessageW(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Tracked Changes");
            }
            SendMessageW(hCombo, CB_SETCURSEL, 0, 0);
            // Password controls initially hidden (first selection is "No protection")
            ShowWindow(GetDlgItem(hDlg, PDL_PASSLBL), SW_HIDE);
            ShowWindow(GetDlgItem(hDlg, PDL_PASS), SW_HIDE);
            return TRUE;
        }
        case WM_COMMAND:
        {
            if (HIWORD(wParam) == CBN_SELCHANGE && LOWORD(wParam) == PDL_COMBO) {
                int sel = (int)SendMessageW(GetDlgItem(hDlg, PDL_COMBO), CB_GETCURSEL, 0, 0);
                bool needPass = (sel > 0);
                ShowWindow(GetDlgItem(hDlg, PDL_PASSLBL), needPass ? SW_SHOW : SW_HIDE);
                ShowWindow(GetDlgItem(hDlg, PDL_PASS), needPass ? SW_SHOW : SW_HIDE);
                return TRUE;
            }
            if (LOWORD(wParam) == IDOK) {
                ProtDlgData* d = reinterpret_cast<ProtDlgData*>(GetWindowLongPtr(hDlg, GWLP_USERDATA));
                HWND hCombo = GetDlgItem(hDlg, PDL_COMBO);
                int sel = (int)SendMessageW(hCombo, CB_GETCURSEL, 0, 0);
                wchar_t modeBuf[64] = {};
                SendMessageW(hCombo, CB_GETLBTEXT, sel, (LPARAM)modeBuf);
                d->chosenMode = modeBuf;
                if (sel > 0) {
                    wchar_t passBuf[256] = {};
                    GetWindowTextW(GetDlgItem(hDlg, PDL_PASS), passBuf, 256);
                    if (!passBuf[0]) {
                        MessageBoxW(hDlg, L"Please enter a password.", L"Document Protection", MB_OK | MB_ICONWARNING);
                        return TRUE;
                    }
                    d->password = passBuf;
                }
                d->accepted = true;
                EndDialog(hDlg, IDOK);
                return TRUE;
            }
            if (LOWORD(wParam) == IDCANCEL) { EndDialog(hDlg, IDCANCEL); return TRUE; }
            break;
        }
        case WM_CLOSE: EndDialog(hDlg, IDCANCEL); return TRUE;
        }
        return FALSE;
    }

    static DLGTEMPLATE* BuildProtDlgTemplate()
    {
        static WORD buf[2048];
        memset(buf, 0, sizeof(buf));
        WORD* p = buf;
        auto W = [&](WORD w) { *p++ = w; };
        auto DW = [&](DWORD dw) { memcpy(p, &dw, 4); p += 2; };
        auto WS = [&](const wchar_t* s) { while (*s) *p++ = (WORD)*s++; *p++ = 0; };
        auto Al = [&]() { if ((DWORD_PTR)p & 2) *p++ = 0; };

        // 6 controls, 240x105 DLUs
        DW(DS_SETFONT | DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU);
        DW(0); W(6); W(0); W(0); W(240); W(105);
        W(0); W(0); WS(L"Document Protection"); W(8); WS(L"MS Shell Dlg");

        auto Item = [&](DWORD style, short x, short y, short w, short h, WORD id,
            const wchar_t* cls, const wchar_t* text) {
                Al();
                DW(style | WS_CHILD | WS_VISIBLE); DW(0);
                W((WORD)x); W((WORD)y); W((WORD)w); W((WORD)h); W(id);
                WS(cls); WS(text); W(0);
            };

        Item(SS_LEFT, 7, 7, 226, 8, 0xFFFF, L"STATIC", L"Protection mode:");
        Item(CBS_DROPDOWNLIST | WS_VSCROLL, 7, 18, 226, 80, PDL_COMBO, L"COMBOBOX", L"");
        Item(SS_LEFT, 7, 50, 226, 8, PDL_PASSLBL, L"STATIC", L"Password (leave blank for no password):");
        Item(ES_AUTOHSCROLL | ES_PASSWORD | WS_BORDER, 7, 61, 226, 14, PDL_PASS, L"EDIT", L"");
        Item(BS_DEFPUSHBUTTON, 130, 83, 50, 14, IDOK, L"BUTTON", L"OK");
        Item(BS_PUSHBUTTON, 183, 83, 50, 14, IDCANCEL, L"BUTTON", L"Cancel");

        return reinterpret_cast<DLGTEMPLATE*>(buf);
    }

    // ----------------------------------------------------------------
    // ContentEditValueW — custom dialog for FIELD_AUTHORS and
    // FIELD_DOCUMENT_PROTECTION.  The three ft_multiplechoice fields
    // (AUTO_UPDATE_STYLES, ANONYMISED_FILES, TCS_ON_OFF) don't need
    // this — TC shows its own dropdown for those automatically.
    // ----------------------------------------------------------------
    __declspec(dllexport) int __stdcall ContentEditValueW(
        HWND parentWin, int fieldIndex, int unitIndex,
        int fieldType, void* fieldValue, int maxLen, int flags,
        const char* langIdentifier)
    {
        DbgLog("ContentEditValue called: fieldIndex=%d fieldType=%d\n", fieldIndex, fieldType);
        if (fieldIndex != FIELD_AUTHORS && fieldIndex != FIELD_DOCUMENT_PROTECTION)
            return ft_notsupported;
        if (!fieldValue) return ft_fieldempty;

        // ---- Author rename dialog ----
        if (fieldIndex == FIELD_AUTHORS) {
            AuthorDlgData dlgData;
            if (DialogBoxIndirectParamW(GetModuleHandleW(nullptr), BuildAuthorDlgTemplate(),
                parentWin, AuthorDlgProc, reinterpret_cast<LPARAM>(&dlgData)) != IDOK || !dlgData.accepted)
                return ft_setcancel;

            // Use a printable delimiter "|" instead of \x01 to avoid UI corruption
            std::wstring encoded = dlgData.oldAuthor + L"|" + dlgData.newAuthor;

            if (fieldType == ft_stringw) {
                // maxLen in TC is passed in bytes. Divide by sizeof(WCHAR) for characters.
                wcsncpy_s(static_cast<WCHAR*>(fieldValue), maxLen / sizeof(WCHAR), encoded.c_str(), _TRUNCATE);
            }
            return ft_setsuccess;
        }

        // ---- Document protection dialog ----
        if (fieldIndex == FIELD_DOCUMENT_PROTECTION) {
            ProtDlgData dlgData;
            // Default to "No protection" initially. Safe for bulk editing.
            dlgData.currentlyProtected = false;

            if (DialogBoxIndirectParamW(GetModuleHandleW(nullptr), BuildProtDlgTemplate(),
                parentWin, ProtDlgProc, reinterpret_cast<LPARAM>(&dlgData)) != IDOK || !dlgData.accepted)
                return ft_setcancel;

            std::wstring encoded = dlgData.chosenMode + L"|" + dlgData.password;

            if (fieldType == ft_stringw) {
                wcsncpy_s(static_cast<WCHAR*>(fieldValue), maxLen / sizeof(WCHAR), encoded.c_str(), _TRUNCATE);
            }
            return ft_setsuccess;
        }

        return ft_notsupported;
    }

    // ANSI shim — TC 32/64-bit ALWAYS calls this because ContentEditValueW is not officially supported by TC.
    __declspec(dllexport) int __stdcall ContentEditValue(
        HWND parentWin, int fieldIndex, int unitIndex,
        int fieldType, char* fieldValue, int maxLen, int flags,
        const char* langIdentifier)
    {
        DbgLog("ContentEditValue called: fieldIndex=%d fieldType=%d\n", fieldIndex, fieldType);
        if (fieldIndex != FIELD_AUTHORS && fieldIndex != FIELD_DOCUMENT_PROTECTION)
            return ft_notsupported;

        // Let ContentEditValueW handle the writing. Because fieldType is ft_stringw, 
        // it will cast fieldValue to WCHAR* and write the string directly into TC's buffer.
        return ContentEditValueW(parentWin, fieldIndex, unitIndex, fieldType,
            static_cast<void*>(fieldValue), maxLen, flags, langIdentifier);
    }

    // Returns a priority score based on ISO/IEC 29500 w:settings sequence.
    int GetSettingsElementPriority(const char* name) {
        if (!name) return 9999;
        std::string s(name);
        if (s == "w:zoom") return 30;
        if (s == "w:removePersonalInformation") return 40;
        if (s == "w:removeDateAndTime") return 50;
        if (s == "w:linkStyles") return 260;
        if (s == "w:trackRevisions") return 320;
        if (s == "w:documentProtection") return 350;
        if (s == "w:defaultTabStop") return 390;
        if (s == "w:compat") return 720;
        if (s == "w:rsids") return 740;
        if (s == "w:mathPr") return 750;
        if (s == "w:themeFontLang") return 770;
        if (s == "w:clrSchemeMapping") return 780;
        if (s == "w:shapeDefaults") return 790;
        if (s == "w:decimalSymbol") return 870;
        if (s == "w:listSeparator") return 880;
        return 999; // Unknowns go near the end, but before core bottom elements
    }

    // Inserts the new element safely before the first element that has a lower priority rank.
    void InsertSettingsElement(tinyxml2::XMLElement* root, tinyxml2::XMLElement* newEl) {
        if (!root || !newEl) return;
        int newRank = GetSettingsElementPriority(newEl->Name());

        for (tinyxml2::XMLElement* child = root->FirstChildElement(); child; child = child->NextSiblingElement()) {
            if (GetSettingsElementPriority(child->Name()) > newRank) {
                // tinyxml2 doesn't have InsertBeforeChild, so we use InsertAfterChild on the previous node
                tinyxml2::XMLNode* prev = child->PreviousSibling();
                if (prev) {
                    root->InsertAfterChild(prev, newEl);
                }
                else {
                    root->InsertFirstChild(newEl); // It's the first element
                }
                return;
            }
        }
        root->InsertEndChild(newEl);
    }

    __declspec(dllexport) int __stdcall ContentSetValueW(
        WCHAR* fileName, int fieldIndex, int unitIndex,
        int fieldType, void* fieldValue, int flags)
    {
        DbgLog("ContentSetValueW called: fieldIndex=%d fieldType=%d fieldValue=%p\n", fieldIndex, fieldType, fieldValue);

        // TC sends fieldIndex=-1 with FileName=NULL as a "commit/finalise" sentinel
        // after all changes for a file. This MUST be checked before the NULL filename
        // guard, because TC passes NULL for FileName on the sentinel call.
        if (fieldIndex == -1) {
            ClearCache();
            return ft_setsuccess;
        }

        if (!fileName) return ft_fieldempty;
        std::string ansiPath;
        if (!WidePathToAnsi(fileName, ansiPath)) return ft_fieldempty;

        g_cancelRequested.store(false, std::memory_order_relaxed);

        auto FieldValueToUtf8 = [&](int fType, const void* fVal) -> std::string {
            if (!fVal) return std::string();
            if (fType == ft_string || fType == ft_stringw || fType == ft_fulltext || fType == ft_fulltextw) {
                const wchar_t* w = static_cast<const wchar_t*>(fVal);
                if (!w) return std::string();
                std::wstring ws(w);
                if (ws.empty()) return std::string();
                int needed = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), -1, NULL, 0, NULL, NULL);
                if (needed <= 0) return std::string();
                std::string s; s.resize(needed - 1);
                WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), -1, &s[0], needed, NULL, NULL);
                return s;
            }
            return std::string();
            };

        switch (fieldIndex) {
            // ----------------------------------------------------------------
            // String fields: Title, Subject, Author, Keywords, Comments (core)
            //                Manager, Company, Hyperlink base (app)
            // ----------------------------------------------------------------
        case FIELD_CORE_TITLE:
        case FIELD_CORE_SUBJECT:
        case FIELD_CORE_CREATOR:
        case FIELD_CORE_KEYWORDS:
        case FIELD_CORE_DESCRIPTION:
        case FIELD_APP_MANAGER:
        case FIELD_APP_COMPANY:
        case FIELD_APP_HYPERLINK_BASE:
        {
            std::string value = FieldValueToUtf8(fieldType, fieldValue);
            std::string currentValue = GetFieldResult(ansiPath, fieldIndex, unitIndex).s;
            if (value == currentValue) return ft_setsuccess;

            bool isCore = (fieldIndex == FIELD_CORE_TITLE ||
                fieldIndex == FIELD_CORE_SUBJECT ||
                fieldIndex == FIELD_CORE_CREATOR ||
                fieldIndex == FIELD_CORE_KEYWORDS ||
                fieldIndex == FIELD_CORE_DESCRIPTION);

            if (isCore) {
                std::string coreXml;
                if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                    const char* elemName =
                        fieldIndex == FIELD_CORE_TITLE ? "dc:title" :
                        fieldIndex == FIELD_CORE_SUBJECT ? "dc:subject" :
                        fieldIndex == FIELD_CORE_CREATOR ? "dc:creator" :
                        fieldIndex == FIELD_CORE_KEYWORDS ? "cp:keywords" :
                        fieldIndex == FIELD_CORE_DESCRIPTION ? "dc:description" : "";
                    if (SetXmlStringValue(coreXml, elemName, value)) {
                        SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml);
                        ClearCache();
                    }
                }
            }
            else {
                std::string appXml;
                if (ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
                    const char* elemName =
                        fieldIndex == FIELD_APP_MANAGER ? "Manager" :
                        fieldIndex == FIELD_APP_COMPANY ? "Company" :
                        fieldIndex == FIELD_APP_HYPERLINK_BASE ? "HyperlinkBase" : "";
                    if (SetXmlStringValue(appXml, elemName, value)) {
                        SaveXmlToZip(ansiPath.c_str(), "docProps/app.xml", appXml);
                        ClearCache();
                    }
                }
            }
        }
        return ft_setsuccess;

        // ----------------------------------------------------------------
        // Date/time fields: Created, Modified, Printed
        // ----------------------------------------------------------------
        case FIELD_CORE_CREATED_DATE:
        case FIELD_CORE_MODIFIED_DATE:
        case FIELD_CORE_LAST_PRINTED_DATE:
        {
            if (fieldType != ft_datetime || !fieldValue) return ft_fieldempty;
            FILETIME ft;
            memcpy(&ft, fieldValue, sizeof(FILETIME));
            std::string newValue = FileTimeToIso8601UTC(&ft);
            if (newValue.empty()) return ft_fieldempty;

            std::string coreXml;
            if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                const char* elemName =
                    fieldIndex == FIELD_CORE_CREATED_DATE ? "dcterms:created" :
                    fieldIndex == FIELD_CORE_MODIFIED_DATE ? "dcterms:modified" :
                    "cp:lastPrinted";
                if (SetXmlStringValue(coreXml, elemName, newValue)) {
                    SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml);
                    ClearCache();
                }
            }
        }
        return ft_setsuccess;

        // ----------------------------------------------------------------
        // Last saved by (core string)
        // ----------------------------------------------------------------
        case FIELD_CORE_LAST_MODIFIED_BY:
        {
            std::string userName = FieldValueToUtf8(fieldType, fieldValue);
            std::string coreXml;
            if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                if (SetXmlStringValue(coreXml, "cp:lastModifiedBy", userName)) {
                    SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml);
                    ClearCache();
                }
            }
        }
        return ft_setsuccess;

        // ----------------------------------------------------------------
        // Revision number (core numeric)
        // ----------------------------------------------------------------
        case FIELD_CORE_REVISION_NUMBER:
        {
            if (!fieldValue) return ft_fieldempty;
            int revision = *(static_cast<const int*>(fieldValue));
            std::string coreXml;
            if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                char buf[16];
                snprintf(buf, sizeof(buf), "%d", revision);
                if (SetXmlStringValue(coreXml, "cp:revision", buf)) {
                    SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml);
                    ClearCache();
                }
            }
        }
        return ft_setsuccess;

        // ----------------------------------------------------------------
        // Total editing time (app numeric, minutes)
        // ----------------------------------------------------------------
        case FIELD_APP_EDITING_TIME:
        {
            if (!fieldValue) return ft_fieldempty;
            int editingTime = *(static_cast<const int*>(fieldValue));
            std::string appXml;
            if (ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
                char buf[16];
                snprintf(buf, sizeof(buf), "%d", editingTime);
                if (SetXmlStringValue(appXml, "TotalTime", buf)) {
                    SaveXmlToZip(ansiPath.c_str(), "docProps/app.xml", appXml);
                    ClearCache();
                }
            }
        }
        return ft_setsuccess;

        // ----------------------------------------------------------------
                // Auto Update Styles — Yes/No -> w:linkStyles
                // ----------------------------------------------------------------
        case FIELD_AUTO_UPDATE_STYLES: {
            if (!fieldValue) return ft_fieldempty;

            // TC passes a 1-based integer index for ft_multiplechoice.
            // Units: "Yes|No" → 1 = Yes (enable), 2 = No (disable)
            const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
            DbgLog("FIELD_AUTO_UPDATE_STYLES: fieldValue=%p nestedChoice=%s\n",
                fieldValue, choiceText ? choiceText : "(null)");
            int idx = NormalizeChoiceIndex(fieldIndex, fieldValue, 2);
            if (idx < 0) return ft_fieldempty;
            bool enable = (idx == 0);
            DbgLog("FIELD_AUTO_UPDATE_STYLES: idx=%d enable=%d\n", idx, (int)enable);

            std::string settingsXml;
            if (!ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml))
                return ft_fileerror;

            tinyxml2::XMLDocument doc;
            if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS)
                return ft_fileerror;

            tinyxml2::XMLElement* root = doc.FirstChildElement("w:settings");
            if (!root) return ft_fileerror;

            tinyxml2::XMLElement* el = root->FirstChildElement("w:linkStyles");
            if (enable) {
                if (!el) {
                    el = doc.NewElement("w:linkStyles");
                    InsertSettingsElement(root, el);
                }
                el->SetAttribute("w:val", "true");
            }
            else {
                if (el) root->DeleteChild(el);
            }

            tinyxml2::XMLPrinter printer(nullptr, true);
            doc.Accept(&printer);
            bool saved = SaveXmlToZip(ansiPath.c_str(), "word/settings.xml", printer.CStr());

            DbgLog("FIELD_AUTO_UPDATE_STYLES: idx=%d enable=%d saved=%d\n", idx, (int)enable, (int)saved);

            ClearCache();
            return ft_setsuccess;
        }

                                     // ----------------------------------------------------------------
                                     // Files Anonymised 
                                     // ----------------------------------------------------------------
        case FIELD_ANONYMISED_FILES: {
            if (!fieldValue) return ft_fieldempty;

            // TC passes a 1-based integer index for ft_multiplechoice.
            // Units: "No|Personal Information|Date and Time|Personal Information and Date and Time"
            // 1 = No, 2 = Personal Information, 3 = Date and Time, 4 = Both
            const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
            DbgLog("FIELD_ANONYMISED_FILES: fieldValue=%p nestedChoice=%s\n",
                fieldValue, choiceText ? choiceText : "(null)");
            int idx = NormalizeChoiceIndex(fieldIndex, fieldValue, 4);
            if (idx < 0) return ft_fieldempty;
            DbgLog("FIELD_ANONYMISED_FILES: idx=%d\n", idx);

            bool removePI = (idx == 1 || idx == 3);
            bool removeDate = (idx == 2 || idx == 3);

            std::string settingsXml;
            if (!ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml))
                return ft_fileerror;

            tinyxml2::XMLDocument doc;
            if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS)
                return ft_fileerror;

            tinyxml2::XMLElement* root = doc.FirstChildElement("w:settings");
            if (!root) return ft_fileerror;

            auto setFlag = [&](const char* localName, bool enableFlag) {
                // Word XML tags are usually prefixed with w:
                std::string fullName = std::string("w:") + localName;

                // Search for both prefixed and non-prefixed just in case
                tinyxml2::XMLElement* el = root->FirstChildElement(fullName.c_str());
                if (!el) el = root->FirstChildElement(localName);

                if (enableFlag) {
                    if (!el) {
                        el = doc.NewElement(fullName.c_str());
                        // Make sure your InsertSettingsElement places it inside <w:settings>
                        InsertSettingsElement(root, el);
                    }
                    el->SetAttribute("w:val", "true");
                }
                else {
                    if (el) {
                        root->DeleteChild(el);
                    }
                }
                };

            // Call without the "w:" prefix now
            setFlag("removePersonalInformation", removePI);
            setFlag("removeDateAndTime", removeDate);

            tinyxml2::XMLPrinter printer(nullptr, true);
            doc.Accept(&printer);
            bool saved = SaveXmlToZip(ansiPath.c_str(), "word/settings.xml", printer.CStr());

            DbgLog("FIELD_ANONYMISED_FILES: idx=%d removePI=%d removeDate=%d saved=%d\n",
                idx, (int)removePI, (int)removeDate, (int)saved);

            ClearCache();
            return ft_setsuccess;
        }

        // ----------------------------------------------------------------
        // Track Changes — Activated/Deactivated
        // ----------------------------------------------------------------
        case FIELD_TCS_ON_OFF: {
            if (!fieldValue) return ft_fieldempty;

            // TC passes a 1-based integer index for ft_multiplechoice.
            // Units: "Activated|Deactivated" → 1 = Activated (enable), 2 = Deactivated (disable)
            const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
            DbgLog("FIELD_TCS_ON_OFF: fieldValue=%p nestedChoice=%s\n",
                fieldValue, choiceText ? choiceText : "(null)");
            int idx = NormalizeChoiceIndex(fieldIndex, fieldValue, 2);
            if (idx < 0) return ft_fieldempty;
            bool enable = (idx == 0);
            DbgLog("FIELD_TCS_ON_OFF: idx=%d enable=%d\n", idx, (int)enable);

            std::string settingsXml;
            if (!ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml))
                return ft_fileerror;

            tinyxml2::XMLDocument doc;
            if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS)
                return ft_fileerror;

            tinyxml2::XMLElement* root = doc.FirstChildElement("w:settings");
            if (!root) return ft_fileerror;

            // Track Changes in Word XML is the <w:trackRevisions> element
            tinyxml2::XMLElement* el = root->FirstChildElement("w:trackRevisions");
            if (enable) {
                if (!el) {
                    el = doc.NewElement("w:trackRevisions");
                    InsertSettingsElement(root, el);
                }
                // Word usually treats the presence of the tag as "on", 
                // but w:val="true" is the explicit standard.
                el->SetAttribute("w:val", "true");
            }
            else {
                if (el) root->DeleteChild(el);
            }

            tinyxml2::XMLPrinter printer(nullptr, true);
            doc.Accept(&printer);
            bool saved = SaveXmlToZip(ansiPath.c_str(), "word/settings.xml", printer.CStr());

            DbgLog("FIELD_TCS_ON_OFF: idx=%d enable=%d saved=%d\n", idx, (int)enable, (int)saved);

            ClearCache();
            return ft_setsuccess;
        }

        // ----------------------------------------------------------------
        // Document Protection — set or clear w:documentProtection
        // ----------------------------------------------------------------
        // ----------------------------------------------------------------
        // Document Protection — value encoded as "Mode\x01Password"
        // from ContentEditValueW.  Always clears enforcement first,
        // then re-applies the chosen mode (+ password hash if given).
        // Uses the Word 97 XOR password hash (w:cryptAlgorithmSid="1")
        // which Word accepts without needing SHA512/bcrypt complexity.
        // ----------------------------------------------------------------
        case FIELD_DOCUMENT_PROTECTION:
        {
            if (!fieldValue) return ft_fieldempty;
            std::wstring encoded(static_cast<const wchar_t*>(fieldValue));

            std::string mode;
            std::wstring wPass;
            auto sep = encoded.find(L'|');
            if (sep == std::wstring::npos) {
                mode = WideToUtf8(encoded);
                wPass = L""; // Fallback to no password if typed manually
            }
            else {
                mode = WideToUtf8(encoded.substr(0, sep));
                wPass = encoded.substr(sep + 1);
            }

            std::string settingsXml;
            if (!ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml)) return ft_fileerror;
            tinyxml2::XMLDocument doc;
            if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS) return ft_fileerror;
            tinyxml2::XMLElement* root = doc.FirstChildElement("w:settings");
            if (!root) return ft_fileerror;

            // Always remove existing protection element first (clears any prior enforcement/password)
            tinyxml2::XMLElement* prot = root->FirstChildElement("w:documentProtection");
            if (prot) root->DeleteChild(prot);

            if (mode != "No protection") {
                const char* editVal =
                    mode == "Read-Only" ? "readOnly" :
                    mode == "Forms" ? "forms" :
                    mode == "Comments" ? "comments" :
                    mode == "Tracked Changes" ? "trackedChanges" : nullptr;
                if (!editVal) return ft_fileerror;

                prot = doc.NewElement("w:documentProtection");
                prot->SetAttribute("w:edit", editVal);
                prot->SetAttribute("w:enforcement", "1");

                if (!wPass.empty()) {
                    // Word 97 XOR password hash — universally accepted by Word versions
                    // Reference: ECMA-376 Part 4 §14.7.1 / MS-OI29500
                    static const WORD encMatrix[15][7] = {
                        {0xAEFC,0x4DD9,0x9BB2,0x2745,0x4E8A,0x9D14,0x2A28},
                        {0x5450,0xA8A0,0x5141,0xA282,0x4584,0x8B08,0x1610},
                        {0x2C20,0x5840,0xB080,0x60B1,0xC162,0x8E45,0x1C8A},
                        {0x3914,0x7228,0xE450,0xCEA1,0x9C43,0x3887,0x710E},
                        {0xE21C,0xC439,0x8873,0x10E6,0x21CC,0x4398,0x8730},
                        {0x0E60,0x1CC0,0x3980,0x7300,0xE600,0xCC01,0x9803},
                        {0x3006,0x600C,0xC018,0x8031,0x0162,0x02C4,0x0588},
                        {0x0B10,0x1620,0x2C40,0x5880,0xB100,0x7201,0xE402},
                        {0xC805,0x8B0B,0x1616,0x2C2C,0x5858,0xB0B0,0x6161},
                        {0xC2C2,0x8585,0x0B0B,0x1616,0x2C2C,0x5858,0xB0B0},
                        {0x6161,0xC2C2,0x8585,0x0B0B,0x1616,0x2C2C,0x5858},
                        {0xB0B0,0x6161,0xC2C2,0x8585,0x0B0B,0x1616,0x2C2C},
                        {0x5858,0xB0B0,0x6161,0xC2C2,0x8585,0x0B0B,0x1616},
                        {0x2C2C,0x5858,0xB0B0,0x6161,0xC2C2,0x8585,0x0B0B},
                        {0x1616,0x2C2C,0x5858,0xB0B0,0x6161,0xC2C2,0x8585}
                    };
                    // Convert password to Latin-1 bytes (Word 97 hashes byte values)
                    std::string pass;
                    for (wchar_t ch : wPass) pass += (char)(ch & 0xFF);
                    int len = (int)pass.size();
                    WORD hash = 0;
                    for (int i = len - 1; i >= 0; --i) {
                        hash ^= (WORD)(unsigned char)pass[i];
                        for (int bit = 0; bit < 7; ++bit) {
                            if (hash & 0x4000) hash = ((hash << 1) & 0x7FFF) ^ 0x6072;
                            else               hash = (hash << 1) & 0x7FFF;
                        }
                    }
                    hash ^= (WORD)len;
                    hash ^= 0xCE4B;
                    char hashHex[8]; snprintf(hashHex, sizeof(hashHex), "%04X", (unsigned)hash);
                    prot->SetAttribute("w:cryptAlgorithmClass", "hash");
                    prot->SetAttribute("w:cryptAlgorithmType", "typeAny");
                    prot->SetAttribute("w:cryptAlgorithmSid", "1");     // XOR (Word 97)
                    prot->SetAttribute("w:cryptSpinCount", "0");
                    prot->SetAttribute("w:hash", hashHex);
                    prot->SetAttribute("w:salt", "");
                }
                InsertSettingsElement(root, prot);
            }

            tinyxml2::XMLPrinter printer(nullptr, true); doc.Accept(&printer);
            SaveXmlToZip(ansiPath.c_str(), "word/settings.xml", printer.CStr());
            ClearCache();
        }
        return ft_setsuccess;

        // ----------------------------------------------------------------
        // Tracked Change Authors — rename; value encoded as "OLD\x01NEW"
        // (written by ContentEditValueW; empty OLD = rename all authors)
        // ----------------------------------------------------------------
        case FIELD_AUTHORS:
        {
            if (!fieldValue) return ft_fieldempty;
            std::wstring encoded(static_cast<const wchar_t*>(fieldValue));

            std::string oldName;
            std::string newName;
            auto sep = encoded.find(L'|');
            if (sep == std::wstring::npos) {
                oldName = ""; // Empty = rename all
                newName = WideToUtf8(encoded);
            }
            else {
                oldName = WideToUtf8(encoded.substr(0, sep));
                newName = WideToUtf8(encoded.substr(sep + 1));
            }

            if (newName.empty()) return ft_fieldempty;

            if (!RenameTrackedChangeAuthors(ansiPath, oldName, newName)) return ft_fileerror;
        }
        return ft_setsuccess;

        default:
            return ft_notsupported;
        }
    }

}

extern "C" __declspec(dllexport) int __stdcall ContentSetValue(
    char* fileName, int fieldIndex, int unitIndex,
    int fieldType, void* fieldValue, int flags)
{
    DbgLog("ContentSetValue called: fieldIndex=%d\n", fieldIndex);

    // TC sends fieldIndex=-1 with FileName=NULL as a "commit/finalise" sentinel.
    // This MUST be checked before the NULL filename guard.
    if (fieldIndex == -1) {
        ClearCache();
        return ft_setsuccess;
    }

    if (!fileName) return ft_fieldempty;

    int needed = MultiByteToWideChar(CP_ACP, 0, fileName, -1, NULL, 0);
    if (needed <= 0) return ft_fieldempty;
    std::wstring wfn; wfn.resize(needed - 1);
    MultiByteToWideChar(CP_ACP, 0, fileName, -1, &wfn[0], needed);

    const void* passValue = fieldValue;
    std::wstring widebuf;

    if (fieldValue) {
        // For string fields only: convert ANSI -> wide.
        // ft_multiplechoice is excluded: TC passes a pointer to a 1-based int index,
        // not a string. Pass the int* straight through unchanged.
        if (fieldType == ft_string || fieldType == ft_stringw) {
            const char* src = static_cast<const char*>(fieldValue);
            int wneeded = MultiByteToWideChar(CP_ACP, 0, src, -1, NULL, 0);
            if (wneeded > 0) {
                widebuf.resize(wneeded - 1);
                MultiByteToWideChar(CP_ACP, 0, src, -1, &widebuf[0], wneeded);
                passValue = widebuf.c_str();
            }
        }
    }

    return ContentSetValueW(&wfn[0], fieldIndex, unitIndex, fieldType, const_cast<void*>(passValue), flags);
}
