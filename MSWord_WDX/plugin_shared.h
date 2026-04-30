#pragma once

#include <windows.h>

#include <atomic>
#include <cstdarg>
#include <cstdint>
#include <map>
#include <mutex>
#include <set>
#include <string>
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

enum FieldIndex {
    FIELD_CORE_TITLE = 0,
    FIELD_CORE_SUBJECT,
    FIELD_CORE_CREATOR,
    FIELD_APP_MANAGER,
    FIELD_APP_COMPANY,
    FIELD_CORE_KEYWORDS,
    FIELD_CORE_DESCRIPTION,
    FIELD_APP_HYPERLINK_BASE,
    FIELD_APP_TEMPLATE,
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
    FIELD_COMPATMODE,
    FIELD_HIDDEN_TEXT,
    FIELD_COMMENTS,
    FIELD_DOCUMENT_PROTECTION,
    FIELD_AUTO_UPDATE_STYLES,
    FIELD_ANONYMISATION,
    FIELD_TRACKED_CHANGES,
    FIELD_TRACK_CHANGES_ENABLED_DISABLED,
    FIELD_AUTHORS,
    FIELD_TOTAL_REVISIONS,
    FIELD_TOTAL_INSERTIONS,
    FIELD_TOTAL_DELETIONS,
    FIELD_TOTAL_MOVES,
    FIELD_TOTAL_FORMATTING_CHANGES,
    FIELD_COUNT
};

struct CachedParts {
    std::string path;
    std::string coreXml;
    std::string appXml;
    std::string commentsXml;
    std::string settingsXml;
    std::string documentXml;
    bool derivedReady = false;
    std::string coreTitle;
    std::string coreSubject;
    std::string coreCreator;
    std::string appManager;
    std::string appCompany;
    std::string coreKeywords;
    std::string coreDescription;
    std::string appHyperlinkBase;
    std::string appTemplate;
    std::string coreCreated;
    std::string coreModified;
    std::string coreLastPrinted;
    std::string coreLastModifiedBy;
    int coreRevision = 0;
    int appEditingTime = 0;
    int appPages = 0;
    int appParagraphs = 0;
    int appLines = 0;
    int appWords = 0;
    int appCharacters = 0;
    bool compatibilityMode = false;
    bool autoUpdateStyles = false;
    int anonymisationFlags = 0;
    bool trackChangesEnabled = false;
    int commentsCount = 0;
    bool hiddenText = false;
};

struct TrackedChangeCounts {
    int insertions = 0;
    int deletions = 0;
    int moves = 0;
    int formattingChanges = 0;
    int totalRevisions = 0;
    std::set<std::string> uniqueFormattingChanges;
};

struct FieldResult {
    enum Type { Empty, FileTime, Int32, Int64, Boolean, StringUtf8 } type = Empty;
    FILETIME ft{};
    int32_t i32 = 0;
    int64_t i64 = 0;
    bool b = false;
    std::string s;
};

struct FieldDescriptor {
    const char* name;
    const char* units;
    int type;
    int flags;
};

struct AuthorRenameEntry {
    std::string oldAuthor;
    std::string newAuthor;
};

extern std::atomic<bool> g_cancelRequested;

const FieldDescriptor* GetFieldDescriptor(int fieldIndex);
const char* GetMultipleChoiceText(int fieldIndex, int value);

bool EnsureCachedParts(const std::string& ansiPath);
void ClearCache();
bool IsCanceled();

bool ExtractFileFromZip(const char* zipPath, const char* fileNameInZip, std::string& output);
bool WidePathToAnsi(const WCHAR* wpath, std::string& out);
bool Utf8ToWideString(const std::string& in, std::wstring& out);
std::string WideToUtf8(const std::wstring& wstr);
std::string FileTimeToIso8601UTC(const FILETIME* ftUtc);
bool ParseIso8601ToFileTime(const std::string& iso8601Str, FILETIME* ftOut);
bool FormatSystemTimeToString(const FILETIME& ftUtc, int unitIndex, wchar_t* outputWStr, int maxWLen);
bool FormatSystemTimeToAnsi(const FILETIME& ftUtc, int unitIndex, char* outputStr, int maxLen);

std::set<std::string> GetTrackedChangeAuthorsFromAllXml(const char* zipPath);
bool HasTrackedChanges(const char* zipPath);
int CountComments(const std::string& xmlContent);
bool IsTrackChangesEnabled(const std::string& settingsXmlContent);
bool IsAutoUpdateStylesEnabled(const std::string& settingsXmlContent);
int GetAnonymisedFlags(const std::string& settingsXmlContent);
bool HasHiddenTextInDocumentXml(const char* zipPath);
bool IsCompatibilityModeEnabled(const std::string& settingsXmlContent);
std::string GetXmlStringValue(const std::string& xmlContent, const char* elementName);
int GetXmlIntValue(const std::string& xmlContent, const char* elementName);
TrackedChangeCounts GetTrackedChangeCounts(const char* zipPath);
FieldResult GetFieldResult(const std::string& ansiPath, int fieldIndex, int unitIndex);

void DbgLog(const char* fmt, ...);
int NormalizeChoiceIndex(int fieldIndex, const void* fieldValue, int choiceCount);
const char* GetIndirectAnsiChoiceText(const void* fieldValue);
void SetPluginDefaultIniPath(const std::string& iniPath);
std::string TranslatePluginText(const std::string& englishText);
void RefreshPluginLanguageCache();

int GetContentValueWInternal(const std::string& ansiPath, int fieldIndex, int unitIndex, void* fieldValue, int maxLen);
int GetContentValueInternal(const std::string& ansiPath, int fieldIndex, int unitIndex, void* fieldValue, int maxLen);

bool SetXmlStringValue(std::string& xmlContent, const char* elementName, const std::string& value);
bool SaveXmlToZip(const char* zipPath, const char* fileNameInZip, const std::string& content);
bool RenameTrackedChangeAuthors(const std::string& ansiPath, const std::string& oldAuthor, const std::string& newAuthor);
bool RenameTrackedChangeAuthorsBatch(const std::string& ansiPath, const std::vector<AuthorRenameEntry>& renames, const std::string& replaceRemainingWith);

int RunContentSetValueW(WCHAR* fileName, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int flags);
int RunContentSetValue(char* fileName, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int flags);

int RunContentEditValueW(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int maxLen, int flags, const char* langIdentifier);
int RunContentEditValue(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, char* fieldValue, int maxLen, int flags, const char* langIdentifier);
