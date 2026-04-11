#include <windows.h>
#include <string>
#include <cstring>
#include <set>
#include <sstream>
#include <functional>
#include <cstddef>
#include <cstdlib>
#include <iomanip>
#include <ctime>
#include <cstdio>
#include <vector>
#include "miniz.h"
#include "tinyxml2.h"

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

typedef struct {
WORD wYear;
WORD wMonth;
WORD wDay;
} tdateformat,*pdateformat;

typedef struct {
WORD wHour;
WORD wMinute;
WORD wSecond;
} ttimeformat,*ptimeformat;

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
            // Some formatting changes might have 'w:originalAuthor' as well
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

bool IsTrackChangesEnabled(const std::string& settingsXmlContent) {
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return false;
    }

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return false;

    for (tinyxml2::XMLElement* child = root->FirstChildElement(); child != nullptr; child = child->NextSiblingElement()) {
        const char* tag = child->Value();
        if (tag && std::string(tag).find("trackRevisions") != std::string::npos) {
            return true;
        }
    }

    return false;
}

bool IsAutoUpdateStylesEnabled(const std::string& settingsXmlContent) {
    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return false;
    }

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return false;

    for (tinyxml2::XMLElement* child = root->FirstChildElement(); child != nullptr; child = child->NextSiblingElement()) {
        const char* tag = child->Value();
        if (tag && std::string(tag).find("linkStyles") != std::string::npos) {
            return true;
        }
    }

    return false;
}

int GetAnonymisedFlags(const std::string& settingsXmlContent) {
    if (settingsXmlContent.empty()) return 0;

    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXmlContent.c_str()) != tinyxml2::XML_SUCCESS) {
        return 0;
    }

    tinyxml2::XMLElement* root = doc.RootElement();
    if (!root) return 0;

    bool hasPI = false;
    bool hasDate = false;

    for (tinyxml2::XMLElement* child = root->FirstChildElement(); child != nullptr; child = child->NextSiblingElement()) {
        const char* tag = child->Value();
        if (!tag) continue;
        std::string t(tag);
        if (t.find("removePersonalInformation") != std::string::npos) {
            hasPI = true;
        }
        if (t.find("removeDateAndTime") != std::string::npos) {
            hasDate = true;
        }
        if (hasPI && hasDate) break;
    }

    int flags = 0;
    if (hasPI) flags |= 1;
    if (hasDate) flags |= 2;
    return flags;
}

bool AreFilesAnonymised(const std::string& settingsXmlContent) {
    return GetAnonymisedFlags(settingsXmlContent) != 0;
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
                } else if (sign == '-') {
                    hour += offset_h;
                    minute += offset_m;
                }
            } else {
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
                // For these property changes, include the type of property (e.g., w:b, w:color, w:pStyle)
                // This helps uniquely identify the specific type of formatting change being tracked.
                if (elem->FirstChildElement()) {
                    change_id += ":" + std::string(elem->FirstChildElement()->Name());
                }
                else {
                    // Fallback: if no child, just use the change tag itself as the unique ID.
                    change_id += ":noChild";
                }
            }
            // Removed direct check for w:style here, as it defines, not tracks, unless wrapped by *PrChange.

            if (isFormattingChange) {
                // Add to unique set to count each *distinct* tracked formatting change only once
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
        mz_zip_archive_file_stat file_stat;
        if (!mz_zip_reader_file_stat(&zip_archive, i, &file_stat))
            continue;

        const char* fname = file_stat.m_filename;
        if (!fname) continue;

        // Process only XML files within the "word/" directory (e.g., document.xml, styles.xml, etc.)
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


// --- Total Commander Content Plugin API ---

extern "C" {

    __declspec(dllexport) int __stdcall ContentGetSupportedField(int fieldIndex, char* fieldName, char* units, int maxLen)
    {
        switch (fieldIndex)
        {
        case FIELD_CORE_TITLE:
            strncpy_s(fieldName, maxLen, "Document Title", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_CORE_SUBJECT:
            strncpy_s(fieldName, maxLen, "Subject", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_CORE_CREATOR:
            strncpy_s(fieldName, maxLen, "Author", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_APP_MANAGER:
            strncpy_s(fieldName, maxLen, "Manager", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_APP_COMPANY:
            strncpy_s(fieldName, maxLen, "Company", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_CORE_KEYWORDS:
            strncpy_s(fieldName, maxLen, "Keywords", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_CORE_DESCRIPTION:
            strncpy_s(fieldName, maxLen, "Comments", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_APP_HYPERLINK_BASE:
            strncpy_s(fieldName, maxLen, "Hyperlink base", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_APP_TEMPLATE:
            strncpy_s(fieldName, maxLen, "Template", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;

        // Statistics Fields
        case FIELD_CORE_CREATED_DATE:
            strncpy_s(fieldName, maxLen, "Created", _TRUNCATE);
            strncpy_s(units, maxLen,"", _TRUNCATE);
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
            return ft_string;
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
            return ft_string;
        case FIELD_AUTO_UPDATE_STYLES:
            strncpy_s(fieldName, maxLen, "Auto Update Styles", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_boolean;
        case FIELD_ANONYMISED_FILES:
            strncpy_s(fieldName, maxLen, "Files Anonymised", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
        case FIELD_TCS_ON_OFF:
            strncpy_s(fieldName, maxLen, "Track Changes", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_boolean;
        case FIELD_TRACKED_CHANGES:
            strncpy_s(fieldName, maxLen, "Tracked Changes Present in Document", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_boolean;
        case FIELD_AUTHORS:
            strncpy_s(fieldName, maxLen, "Tracked Changes Authors", _TRUNCATE);
            strncpy_s(units, maxLen, "", _TRUNCATE);
            return ft_string;
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

    __declspec(dllexport) int __stdcall ContentGetValue(
        char* fileName, int fieldIndex, int unitIndex,
        void* fieldValue, int maxLen, int flags)
    {
        size_t len = strlen(fileName);
        if (len < 5 || _stricmp(fileName + len - 5, ".docx") != 0)
            return ft_fieldempty;

        std::string documentXml;
        std::string commentsXml;
        std::string settingsXml;
        std::string coreXml;
        std::string appXml;

        bool needsCoreXml = false;
        bool needsAppXml = false;
        bool needsWordXml = false;
        bool needsCommentsXml = false;
        bool needsSettingsXml = false;

        // Determine which XML files are needed based on the fieldIndex
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
                needsCoreXml = true;
                break;
            case FIELD_APP_MANAGER:
            case FIELD_APP_COMPANY:
            case FIELD_APP_HYPERLINK_BASE:
            case FIELD_APP_TEMPLATE:
            case FIELD_APP_PAGES:
            case FIELD_APP_WORDS:
            case FIELD_APP_CHARACTERS:
            case FIELD_APP_LINES:
            case FIELD_APP_PARAGRAPHS:
            case FIELD_APP_EDITING_TIME:
                needsAppXml = true;
                break;
            case FIELD_COMPATMODE:
            case FIELD_COMMENTS:
                needsCommentsXml = true;
                break;
            case FIELD_DOCUMENT_PROTECTION:
                needsSettingsXml = true;
                break;
            case FIELD_AUTO_UPDATE_STYLES:
            case FIELD_ANONYMISED_FILES:
            case FIELD_TCS_ON_OFF:
            case FIELD_HIDDEN_TEXT:
            case FIELD_TRACKED_CHANGES:
            case FIELD_TOTAL_REVISIONS:
            case FIELD_TOTAL_INSERTIONS:
            case FIELD_TOTAL_DELETIONS:
            case FIELD_TOTAL_MOVES:
            case FIELD_TOTAL_FORMATTING_CHANGES:
                needsWordXml = true;
                break;
            case FIELD_AUTHORS:
                break;
            default:
                break;
        }

        if (needsCoreXml) {
            if (!ExtractFileFromZip(fileName, "docProps/core.xml", coreXml)) {
                return ft_fileerror;
            }
        }
        if (needsAppXml) {
            if (!ExtractFileFromZip(fileName, "docProps/app.xml", appXml)) {
                return ft_fileerror;
            }
        }
        if (needsCommentsXml) {
            if (!ExtractFileFromZip(fileName, "word/comments.xml", commentsXml)) {
                commentsXml = "";
            }
        }
        if (needsSettingsXml) {
            if (!ExtractFileFromZip(fileName, "word/settings.xml", settingsXml)) {
                return ft_fileerror;
            }
        }
        if (needsWordXml) {
            if (!ExtractFileFromZip(fileName, "word/document.xml", documentXml)) {
                return ft_fileerror;
            }
        }

        TrackedChangeCounts trackedCounts;
        if (fieldIndex >= FIELD_TOTAL_REVISIONS && fieldIndex <= FIELD_TOTAL_FORMATTING_CHANGES) {
            trackedCounts = GetTrackedChangeCounts(fileName); // Call the updated function
        }


        switch (fieldIndex)
        {
        case FIELD_CORE_TITLE:
        {
            std::string title = GetXmlStringValue(coreXml, "dc:title");
            if (title.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, title.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_CORE_SUBJECT:
        {
            std::string subject = GetXmlStringValue(coreXml, "dc:subject");
            if (subject.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, subject.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_CORE_CREATOR:
        {
            std::string creator = GetXmlStringValue(coreXml, "dc:creator");
            if (creator.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, creator.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_APP_MANAGER:
        {
            std::string manager = GetXmlStringValue(appXml, "Manager");
            if (manager.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, manager.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_APP_COMPANY:
        {
            std::string company = GetXmlStringValue(appXml, "Company");
            if (company.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, company.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_CORE_KEYWORDS:
        {
            std::string keywords = GetXmlStringValue(coreXml, "cp:keywords");
            if (keywords.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, keywords.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_CORE_DESCRIPTION:
        {
            std::string description = GetXmlStringValue(coreXml, "dc:description");
            if (description.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, description.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_APP_HYPERLINK_BASE:
        {
            std::string hyperlinkBase = GetXmlStringValue(appXml, "HyperlinkBase");
            if (hyperlinkBase.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, hyperlinkBase.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_APP_TEMPLATE:
        {
            std::string templateName = GetXmlStringValue(appXml, "Template");
            if (templateName.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, templateName.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_CORE_CREATED_DATE:
        case FIELD_CORE_MODIFIED_DATE:
        case FIELD_CORE_LAST_PRINTED_DATE:
        {
            std::string dateStr;
            if (fieldIndex == FIELD_CORE_CREATED_DATE) dateStr = GetXmlStringValue(coreXml, "dcterms:created");
            else if (fieldIndex == FIELD_CORE_MODIFIED_DATE) dateStr = GetXmlStringValue(coreXml, "dcterms:modified");
            else if (fieldIndex == FIELD_CORE_LAST_PRINTED_DATE) dateStr = GetXmlStringValue(coreXml, "cp:lastPrinted");

            if (dateStr.empty()) return ft_fieldempty;

            FILETIME ft_utc;
            if (!ParseIso8601ToFileTime(dateStr, &ft_utc)) {
                return ft_fieldempty;
            }

            if (unitIndex == 0) {
                memcpy(fieldValue, &ft_utc, sizeof(FILETIME));
                return ft_datetime;
            }
            else {
                if (FormatSystemTimeToString(ft_utc, unitIndex, static_cast<wchar_t*>(fieldValue), maxLen / sizeof(wchar_t))) {
                    return ft_stringw;
                }
                return ft_fieldempty;
            }
        }
        case FIELD_CORE_LAST_MODIFIED_BY:
        {
            std::string lastModifiedBy = GetXmlStringValue(coreXml, "cp:lastModifiedBy");
            if (lastModifiedBy.empty()) return ft_fieldempty;
            strncpy_s(static_cast<char*>(fieldValue), maxLen, lastModifiedBy.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_CORE_REVISION_NUMBER:
        {
            int revision = GetXmlIntValue(coreXml, "cp:revision");
            if (revision == 0 && coreXml.empty()) return ft_fileerror;
            *(int*)fieldValue = revision;
            return ft_numeric_32;
        }
        case FIELD_APP_EDITING_TIME:
        {
            int editTime = GetXmlIntValue(appXml, "TotalTime");
            if (appXml.empty()) return ft_fileerror;
            *(int*)fieldValue = editTime;
            return ft_numeric_32;
        }
        case FIELD_APP_PAGES:
        {
            int pages = GetXmlIntValue(appXml, "Pages");
            if (pages == 0 && appXml.empty()) return ft_fileerror;
            if (pages == 0) return ft_fieldempty;
            *(int*)fieldValue = pages;
            return ft_numeric_32;
        }
        case FIELD_APP_PARAGRAPHS:
        {
            int paragraphs = GetXmlIntValue(appXml, "Paragraphs");
            if (paragraphs == 0 && appXml.empty()) return ft_fileerror;
            if (paragraphs == 0) return ft_fieldempty;
            *(int*)fieldValue = paragraphs;
            return ft_numeric_32;
        }
        case FIELD_APP_LINES:
        {
            int lines = GetXmlIntValue(appXml, "Lines");
            if (lines == 0 && appXml.empty()) return ft_fileerror;
            if (lines == 0) return ft_fieldempty;
            *(int*)fieldValue = lines;
            return ft_numeric_32;
        }
        case FIELD_APP_WORDS:
        {
            int words = GetXmlIntValue(appXml, "Words");
            if (words == 0 && appXml.empty()) return ft_fileerror;
            if (words == 0) return ft_fieldempty;
            *(int*)fieldValue = words;
            return ft_numeric_32;
        }
        case FIELD_APP_CHARACTERS:
        {
            int characters = GetXmlIntValue(appXml, "Characters");
            if (characters == 0 && appXml.empty()) return ft_fileerror;
            if (characters == 0) return ft_fieldempty;
            *(int*)fieldValue = characters;
            return ft_numeric_32;
        }

        case FIELD_COMPATMODE:
        {
            if (!ExtractFileFromZip(fileName, "word/settings.xml", settingsXml))
                return ft_fileerror;

            *((int*)fieldValue) = IsCompatibilityModeEnabled(settingsXml) ? 1 : 0;
            return ft_boolean;
        }
        case FIELD_HIDDEN_TEXT:
        {
            bool hasHidden = HasHiddenTextInDocumentXml(fileName);
            *((int*)fieldValue) = hasHidden ? 1 : 0;
            return ft_boolean;
        }
        case FIELD_COMMENTS:
        {
            int commentCount = CountComments(commentsXml);
            *(int*)fieldValue = commentCount;
            return ft_numeric_32;
        }
        case FIELD_DOCUMENT_PROTECTION:
        {
            if (settingsXml.empty()) return ft_fileerror;

            tinyxml2::XMLDocument doc;
            if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS) {
                strncpy_s(static_cast<char*>(fieldValue), maxLen, "Error parsing settings.xml", _TRUNCATE);
                static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
                return ft_string;
            }

            tinyxml2::XMLElement* root = doc.RootElement();
            if (!root) {
                strncpy_s(static_cast<char*>(fieldValue), maxLen, "No protection", _TRUNCATE);
                static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
                return ft_string;
            }

            tinyxml2::XMLElement* protectionElem = root->FirstChildElement("w:documentProtection");
            std::string protectionType = "No protection";

            if (protectionElem) {
                const char* enforcement = protectionElem->Attribute("w:enforcement");
                if (enforcement && strcmp(enforcement, "1") == 0) {
                    const char* edit = protectionElem->Attribute("w:edit");
                    if (edit) {
                        if (strcmp(edit, "readOnly") == 0) {
                            protectionType = "Read-Only";
                        }
                        else if (strcmp(edit, "forms") == 0) {
                            protectionType = "Forms";
                        }
                        else if (strcmp(edit, "comments") == 0) {
                            protectionType = "Comments";
                        }
                        else if (strcmp(edit, "trackedChanges") == 0) {
                            protectionType = "Tracked Changes";
                        }
                        else {
                            protectionType = "Unknown protection type";
                        }
                    }
                    else {
                        protectionType = "Unknown protection type";
                    }
                }
            }

            strncpy_s(static_cast<char*>(fieldValue), maxLen, protectionType.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }

        case FIELD_AUTO_UPDATE_STYLES:
        {
            if (!ExtractFileFromZip(fileName, "word/settings.xml", settingsXml)) {
                return ft_fileerror;
            }
            if (IsAutoUpdateStylesEnabled(settingsXml)) {
                *(int*)fieldValue = 1;
            }
            else {
                *(int*)fieldValue = 0;
            }
            return ft_boolean;
        }
        case FIELD_ANONYMISED_FILES:
        {
            if (!ExtractFileFromZip(fileName, "word/settings.xml", settingsXml)) {
                return ft_fileerror;
            }

            int flags = GetAnonymisedFlags(settingsXml);
            std::string report;
            switch (flags) {
            case 0:
                report = "No";
                break;
            case 1:
                report = "Personal Information";
                break;
            case 2:
                report = "Date and Time";
                break;
            case 3:
                report = "Personal Information, Date and Time";
                break;
            default:
                report = "No";
                break;
            }

            strncpy_s(static_cast<char*>(fieldValue), maxLen, report.c_str(), _TRUNCATE);
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_TRACKED_CHANGES:
        {
            bool hasChanges = HasTrackedChanges(fileName); // Call the updated function
            *((int*)fieldValue) = hasChanges ? 1 : 0;
            return ft_boolean;
        }
        case FIELD_TCS_ON_OFF:
        {
            std::string settingsXml;
            
            if (!ExtractFileFromZip(fileName, "word/settings.xml", settingsXml)) {
                strncpy_s(static_cast<char*>(fieldValue), maxLen, "Deactivated", _TRUNCATE);
                static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
                return ft_string;
            }

            if (IsTrackChangesEnabled(settingsXml)) {
                strncpy_s(static_cast<char*>(fieldValue), maxLen, "Activated", _TRUNCATE);
            }
            else {
                strncpy_s(static_cast<char*>(fieldValue), maxLen, "Deactivated", _TRUNCATE);
            }
            static_cast<char*>(fieldValue)[maxLen - 1] = '\0';
            return ft_string;
        }
        case FIELD_AUTHORS:
        {
            auto authors = GetTrackedChangeAuthorsFromAllXml(fileName);
            if (authors.empty())
                return ft_fieldempty;

            std::stringstream ss;
            bool first = true;
            for (const auto& author : authors)
            {
                if (!first) ss << ", ";
                ss << author;
                first = false;
            }
            std::string result = ss.str();

            char* pStr = static_cast<char*>(fieldValue);
            strncpy_s(pStr, maxLen, result.c_str(), _TRUNCATE);
            pStr[maxLen - 1] = '\0';

            return ft_string;
        }
        case FIELD_TOTAL_REVISIONS:
        {
            *(int*)fieldValue = trackedCounts.totalRevisions;
            return ft_numeric_32;
        }
        case FIELD_TOTAL_INSERTIONS:
        {
            *(int*)fieldValue = trackedCounts.insertions;
            return ft_numeric_32;
        }
        case FIELD_TOTAL_DELETIONS:
        {
            *(int*)fieldValue = trackedCounts.deletions;
            return ft_numeric_32;
        }
        case FIELD_TOTAL_MOVES:
        {
            *(int*)fieldValue = trackedCounts.moves;
            return ft_numeric_32;
        }
        case FIELD_TOTAL_FORMATTING_CHANGES:
        {
            *(int*)fieldValue = trackedCounts.formattingChanges;
            return ft_numeric_32;
        }

        default:
            return ft_nomorefields;
        }
    }

}
