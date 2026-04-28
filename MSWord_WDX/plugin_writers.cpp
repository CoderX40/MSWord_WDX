#include "pch.h"

#include "plugin_shared.h"

#include <algorithm>
#include <functional>
#include <map>
#include <vector>

namespace {

std::vector<std::wstring> SplitLines(const std::wstring& text)
{
    std::vector<std::wstring> lines;
    std::wstring current;
    for (wchar_t ch : text) {
        if (ch == L'\r') {
            continue;
        }
        if (ch == L'\n') {
            lines.push_back(current);
            current.clear();
            continue;
        }
        current.push_back(ch);
    }
    lines.push_back(current);
    return lines;
}

std::vector<std::wstring> SplitTabs(const std::wstring& text)
{
    std::vector<std::wstring> parts;
    std::wstring current;
    for (wchar_t ch : text) {
        if (ch == L'\t') {
            parts.push_back(current);
            current.clear();
            continue;
        }
        current.push_back(ch);
    }
    parts.push_back(current);
    return parts;
}

void ParseAuthorRenamePayload(const std::wstring& encoded, std::vector<AuthorRenameEntry>& renames, std::string& replaceRemainingWith)
{
    renames.clear();
    replaceRemainingWith.clear();

    if (encoded.find(L"PAIR\t") == std::wstring::npos && encoded.find(L"ALL\t") == std::wstring::npos) {
        size_t sep = encoded.find(L'|');
        if (sep == std::wstring::npos) {
            std::string newName = WideToUtf8(encoded);
            if (!newName.empty()) {
                replaceRemainingWith = newName;
            }
            return;
        }

        std::string oldAuthor = WideToUtf8(encoded.substr(0, sep));
        std::string newAuthor = WideToUtf8(encoded.substr(sep + 1));
        if (oldAuthor.empty()) {
            replaceRemainingWith = newAuthor;
        }
        else if (!newAuthor.empty()) {
            renames.push_back({ oldAuthor, newAuthor });
        }
        return;
    }

    for (const std::wstring& line : SplitLines(encoded)) {
        if (line.empty()) continue;
        std::vector<std::wstring> parts = SplitTabs(line);
        if (parts.empty()) continue;

        if (parts[0] == L"PAIR" && parts.size() >= 3) {
            std::string oldAuthor = WideToUtf8(parts[1]);
            std::string newAuthor = WideToUtf8(parts[2]);
            if (!oldAuthor.empty() && !newAuthor.empty()) {
                renames.push_back({ oldAuthor, newAuthor });
            }
        }
        else if (parts[0] == L"ALL" && parts.size() >= 2) {
            replaceRemainingWith = WideToUtf8(parts[1]);
        }
    }
}

bool SaveReplacementsToDocx(const std::string& path, const std::map<std::string, std::string>& replacements)
{
    std::string tmpPath = path + ".tmp";

    mz_zip_archive reader{};
    if (!mz_zip_reader_init_file(&reader, path.c_str(), 0)) return false;

    mz_zip_archive writer{};
    if (!mz_zip_writer_init_file(&writer, tmpPath.c_str(), 0)) {
        mz_zip_reader_end(&reader);
        return false;
    }

    mz_uint numFiles = mz_zip_reader_get_num_files(&reader);
    for (mz_uint i = 0; i < numFiles; ++i) {
        mz_zip_archive_file_stat stat;
        if (!mz_zip_reader_file_stat(&reader, i, &stat)) continue;
        if (!stat.m_filename) continue;

        std::string name(stat.m_filename);
        std::map<std::string, std::string>::const_iterator it = replacements.find(name);
        if (it != replacements.end()) {
            const std::string& data = it->second;
            if (!mz_zip_writer_add_mem(&writer, name.c_str(), data.data(), data.size(), MZ_DEFAULT_COMPRESSION)) {
                mz_zip_writer_end(&writer);
                mz_zip_reader_end(&reader);
                return false;
            }
        }
        else if (!mz_zip_writer_add_from_zip_reader(&writer, &reader, i)) {
            mz_zip_writer_end(&writer);
            mz_zip_reader_end(&reader);
            return false;
        }
    }

    for (const auto& kv : replacements) {
        bool exists = false;
        for (mz_uint i = 0; i < numFiles; ++i) {
            mz_zip_archive_file_stat stat;
            if (!mz_zip_reader_file_stat(&reader, i, &stat)) continue;
            if (stat.m_filename && kv.first == stat.m_filename) {
                exists = true;
                break;
            }
        }

        if (!exists) {
            if (!mz_zip_writer_add_mem(&writer, kv.first.c_str(), kv.second.data(), kv.second.size(), MZ_DEFAULT_COMPRESSION)) {
                mz_zip_writer_end(&writer);
                mz_zip_reader_end(&reader);
                return false;
            }
        }
    }

    if (!mz_zip_writer_finalize_archive(&writer)) {
        mz_zip_writer_end(&writer);
        mz_zip_reader_end(&reader);
        return false;
    }

    mz_zip_writer_end(&writer);
    mz_zip_reader_end(&reader);

    return MoveFileExA(tmpPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED) != FALSE;
}

bool UpdateSettingsXml(const std::string& ansiPath, const std::function<bool(tinyxml2::XMLDocument&, tinyxml2::XMLElement*)>& mutator)
{
    std::string settingsXml;
    if (!ExtractFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml)) return false;

    tinyxml2::XMLDocument doc;
    if (doc.Parse(settingsXml.c_str()) != tinyxml2::XML_SUCCESS) return false;

    tinyxml2::XMLElement* root = doc.FirstChildElement("w:settings");
    if (!root) return false;

    if (!mutator(doc, root)) return false;

    tinyxml2::XMLPrinter printer(nullptr, true);
    doc.Accept(&printer);
    bool saved = SaveXmlToZip(ansiPath.c_str(), "word/settings.xml", printer.CStr());
    if (saved) ClearCache();
    return saved;
}

int GetSettingsElementPriority(const char* name)
{
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
    return 999;
}

void InsertSettingsElement(tinyxml2::XMLElement* root, tinyxml2::XMLElement* newEl)
{
    if (!root || !newEl) return;
    int newRank = GetSettingsElementPriority(newEl->Name());

    for (tinyxml2::XMLElement* child = root->FirstChildElement(); child; child = child->NextSiblingElement()) {
        if (GetSettingsElementPriority(child->Name()) > newRank) {
            tinyxml2::XMLNode* prev = child->PreviousSibling();
            if (prev) root->InsertAfterChild(prev, newEl);
            else root->InsertFirstChild(newEl);
            return;
        }
    }

    root->InsertEndChild(newEl);
}

bool AnsiToWideAcp(const char* src, std::wstring& out)
{
    if (!src) {
        out.clear();
        return false;
    }

    int needed = MultiByteToWideChar(CP_ACP, 0, src, -1, nullptr, 0);
    if (needed <= 0) return false;

    std::wstring buffer(static_cast<size_t>(needed), L'\0');
    if (MultiByteToWideChar(CP_ACP, 0, src, -1, &buffer[0], needed) == 0) return false;

    out.assign(buffer.c_str());
    return true;
}

std::string FieldValueToUtf8(int fieldType, const void* fieldValue)
{
    if (!fieldValue) return std::string();
    if (fieldType == ft_string || fieldType == ft_stringw || fieldType == ft_fulltext || fieldType == ft_fulltextw) {
        const wchar_t* w = static_cast<const wchar_t*>(fieldValue);
        if (!w) return std::string();
        return WideToUtf8(std::wstring(w));
    }
    return std::string();
}

} // namespace

bool SetXmlStringValue(std::string& xmlContent, const char* elementName, const std::string& value)
{
    if (xmlContent.empty() || !elementName || elementName[0] == '\0') return false;

    std::string xmlDecl;
    if (xmlContent.rfind("<?xml", 0) == 0) {
        size_t pos = xmlContent.find("?>");
        if (pos != std::string::npos) xmlDecl = xmlContent.substr(0, pos + 2);
    }

    tinyxml2::XMLDocument doc;
    if (doc.Parse(xmlContent.c_str()) != tinyxml2::XML_SUCCESS) return false;

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
            std::string xmlnsAttr = "xmlns:" + prefix;
            if (!root->Attribute(xmlnsAttr.c_str())) {
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

    if (!xmlDecl.empty() && printed.rfind("<?xml", 0) != 0) xmlContent = xmlDecl + "\n" + printed;
    else xmlContent = printed;

    return true;
}

bool SaveXmlToZip(const char* zipPath, const char* fileNameInZip, const std::string& content)
{
    std::map<std::string, std::string> replacements;
    replacements[fileNameInZip] = content;
    return SaveReplacementsToDocx(std::string(zipPath), replacements);
}

bool RenameTrackedChangeAuthors(const std::string& ansiPath, const std::string& oldAuthor, const std::string& newAuthor)
{
    std::vector<AuthorRenameEntry> renames;
    renames.push_back({ oldAuthor, newAuthor });
    return RenameTrackedChangeAuthorsBatch(ansiPath, renames, std::string());
}

bool RenameTrackedChangeAuthorsBatch(const std::string& ansiPath, const std::vector<AuthorRenameEntry>& renames, const std::string& replaceRemainingWith)
{
    mz_zip_archive reader{};
    if (!mz_zip_reader_init_file(&reader, ansiPath.c_str(), 0)) return false;

    std::map<std::string, std::string> replacements;
    bool hasReplaceRemaining = !replaceRemainingWith.empty();
    mz_uint numFiles = mz_zip_reader_get_num_files(&reader);
    for (mz_uint i = 0; i < numFiles; ++i) {
        mz_zip_archive_file_stat stat;
        if (!mz_zip_reader_file_stat(&reader, i, &stat)) continue;
        if (!stat.m_filename) continue;
        if (strncmp(stat.m_filename, "word/", 5) != 0 || strstr(stat.m_filename, ".xml") == nullptr) continue;

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

                const AuthorRenameEntry* matchedRename = nullptr;
                for (const AuthorRenameEntry& rename : renames) {
                    if (rename.oldAuthor == val) {
                        matchedRename = &rename;
                        break;
                    }
                }

                if (matchedRename) {
                    if (matchedRename->newAuthor != val) {
                        elem->SetAttribute(attr, matchedRename->newAuthor.c_str());
                        modified = true;
                    }
                }
                else if (hasReplaceRemaining && replaceRemainingWith != val) {
                    elem->SetAttribute(attr, replaceRemainingWith.c_str());
                    modified = true;
                }
            }
            for (tinyxml2::XMLElement* child = elem->FirstChildElement(); child; child = child->NextSiblingElement()) {
                walk(child);
            }
        };

        walk(doc.RootElement());

        if (modified) {
            tinyxml2::XMLPrinter printer(nullptr, true);
            doc.Accept(&printer);
            replacements[stat.m_filename] = printer.CStr();
        }
    }

    mz_zip_reader_end(&reader);
    if (replacements.empty()) return true;

    bool saved = SaveReplacementsToDocx(ansiPath, replacements);
    if (saved) ClearCache();
    return saved;
}

int RunContentSetValueW(WCHAR* fileName, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int flags)
{
    UNREFERENCED_PARAMETER(unitIndex);
    UNREFERENCED_PARAMETER(flags);

    DbgLog("ContentSetValueW called: fieldIndex=%d fieldType=%d fieldValue=%p\n", fieldIndex, fieldType, fieldValue);

    if (fieldIndex == -1) {
        ClearCache();
        return ft_setsuccess;
    }

    if (!fileName) return ft_fieldempty;

    std::string ansiPath;
    if (!WidePathToAnsi(fileName, ansiPath)) return ft_fieldempty;

    g_cancelRequested.store(false, std::memory_order_relaxed);

    switch (fieldIndex) {
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

        bool isCore =
            fieldIndex == FIELD_CORE_TITLE ||
            fieldIndex == FIELD_CORE_SUBJECT ||
            fieldIndex == FIELD_CORE_CREATOR ||
            fieldIndex == FIELD_CORE_KEYWORDS ||
            fieldIndex == FIELD_CORE_DESCRIPTION;

        if (isCore) {
            std::string coreXml;
            if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                const char* elemName =
                    fieldIndex == FIELD_CORE_TITLE ? "dc:title" :
                    fieldIndex == FIELD_CORE_SUBJECT ? "dc:subject" :
                    fieldIndex == FIELD_CORE_CREATOR ? "dc:creator" :
                    fieldIndex == FIELD_CORE_KEYWORDS ? "cp:keywords" :
                    "dc:description";

                if (SetXmlStringValue(coreXml, elemName, value) && SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
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
                    "HyperlinkBase";

                if (SetXmlStringValue(appXml, elemName, value) && SaveXmlToZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
                    ClearCache();
                }
            }
        }
        return ft_setsuccess;
    }
    case FIELD_CORE_CREATED_DATE:
    case FIELD_CORE_MODIFIED_DATE:
    case FIELD_CORE_LAST_PRINTED_DATE:
    {
        if (fieldType != ft_datetime || !fieldValue) return ft_fieldempty;
        FILETIME ft{};
        memcpy(&ft, fieldValue, sizeof(FILETIME));
        std::string newValue = FileTimeToIso8601UTC(&ft);
        if (newValue.empty()) return ft_fieldempty;

        std::string coreXml;
        if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            const char* elemName =
                fieldIndex == FIELD_CORE_CREATED_DATE ? "dcterms:created" :
                fieldIndex == FIELD_CORE_MODIFIED_DATE ? "dcterms:modified" :
                "cp:lastPrinted";

            if (SetXmlStringValue(coreXml, elemName, newValue) && SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                ClearCache();
            }
        }
        return ft_setsuccess;
    }
    case FIELD_CORE_LAST_MODIFIED_BY:
    {
        std::string userName = FieldValueToUtf8(fieldType, fieldValue);
        std::string coreXml;
        if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            if (SetXmlStringValue(coreXml, "cp:lastModifiedBy", userName) && SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                ClearCache();
            }
        }
        return ft_setsuccess;
    }
    case FIELD_CORE_REVISION_NUMBER:
    {
        if (!fieldValue) return ft_fieldempty;
        int revision = *static_cast<const int*>(fieldValue);
        std::string coreXml;
        if (ExtractFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            char buf[16];
            snprintf(buf, sizeof(buf), "%d", revision);
            if (SetXmlStringValue(coreXml, "cp:revision", buf) && SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                ClearCache();
            }
        }
        return ft_setsuccess;
    }
    case FIELD_APP_EDITING_TIME:
    {
        if (!fieldValue) return ft_fieldempty;
        int editingTime = *static_cast<const int*>(fieldValue);
        std::string appXml;
        if (ExtractFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
            char buf[16];
            snprintf(buf, sizeof(buf), "%d", editingTime);
            if (SetXmlStringValue(appXml, "TotalTime", buf) && SaveXmlToZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
                ClearCache();
            }
        }
        return ft_setsuccess;
    }
    case FIELD_AUTO_UPDATE_STYLES:
    {
        if (!fieldValue) return ft_fieldempty;
        const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
        DbgLog("FIELD_AUTO_UPDATE_STYLES: fieldValue=%p nestedChoice=%s\n", fieldValue, choiceText ? choiceText : "(null)");
        int idx = NormalizeChoiceIndex(fieldIndex, fieldValue, 2);
        if (idx < 0) return ft_fieldempty;
        bool enable = idx == 0;
        DbgLog("FIELD_AUTO_UPDATE_STYLES: idx=%d enable=%d\n", idx, static_cast<int>(enable));

        bool saved = UpdateSettingsXml(ansiPath, [&](tinyxml2::XMLDocument& doc, tinyxml2::XMLElement* root) {
            tinyxml2::XMLElement* el = root->FirstChildElement("w:linkStyles");
            if (enable) {
                if (!el) {
                    el = doc.NewElement("w:linkStyles");
                    InsertSettingsElement(root, el);
                }
                el->SetAttribute("w:val", "true");
            }
            else if (el) {
                root->DeleteChild(el);
            }
            return true;
        });

        DbgLog("FIELD_AUTO_UPDATE_STYLES: idx=%d enable=%d saved=%d\n", idx, static_cast<int>(enable), static_cast<int>(saved));
        return saved ? ft_setsuccess : ft_fileerror;
    }
    case FIELD_ANONYMISATION:
    {
        if (!fieldValue) return ft_fieldempty;
        const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
        DbgLog("FIELD_ANONYMISATION: fieldValue=%p nestedChoice=%s\n", fieldValue, choiceText ? choiceText : "(null)");
        int idx = NormalizeChoiceIndex(fieldIndex, fieldValue, 4);
        if (idx < 0) return ft_fieldempty;

        bool removePI = idx == 1 || idx == 3;
        bool removeDate = idx == 2 || idx == 3;

        bool saved = UpdateSettingsXml(ansiPath, [&](tinyxml2::XMLDocument& doc, tinyxml2::XMLElement* root) {
            auto setFlag = [&](const char* localName, bool enableFlag) {
                std::string fullName = std::string("w:") + localName;
                tinyxml2::XMLElement* el = root->FirstChildElement(fullName.c_str());
                if (!el) el = root->FirstChildElement(localName);

                if (enableFlag) {
                    if (!el) {
                        el = doc.NewElement(fullName.c_str());
                        InsertSettingsElement(root, el);
                    }
                    el->SetAttribute("w:val", "true");
                }
                else if (el) {
                    root->DeleteChild(el);
                }
            };

            setFlag("removePersonalInformation", removePI);
            setFlag("removeDateAndTime", removeDate);
            return true;
        });

        DbgLog("FIELD_ANONYMISATION: idx=%d removePI=%d removeDate=%d saved=%d\n", idx, static_cast<int>(removePI), static_cast<int>(removeDate), static_cast<int>(saved));
        return saved ? ft_setsuccess : ft_fileerror;
    }
    case FIELD_TRACK_CHANGES_ENABLED_DISABLED:
    {
        if (!fieldValue) return ft_fieldempty;
        const char* choiceText = GetIndirectAnsiChoiceText(fieldValue);
        DbgLog("FIELD_TRACK_CHANGES_ENABLED_DISABLED: fieldValue=%p nestedChoice=%s\n", fieldValue, choiceText ? choiceText : "(null)");
        int idx = NormalizeChoiceIndex(fieldIndex, fieldValue, 2);
        if (idx < 0) return ft_fieldempty;
        bool enable = idx == 0;

        bool saved = UpdateSettingsXml(ansiPath, [&](tinyxml2::XMLDocument& doc, tinyxml2::XMLElement* root) {
            tinyxml2::XMLElement* el = root->FirstChildElement("w:trackRevisions");
            if (enable) {
                if (!el) {
                    el = doc.NewElement("w:trackRevisions");
                    InsertSettingsElement(root, el);
                }
                el->SetAttribute("w:val", "true");
            }
            else if (el) {
                root->DeleteChild(el);
            }
            return true;
        });

        DbgLog("FIELD_TRACK_CHANGES_ENABLED_DISABLED: idx=%d enable=%d saved=%d\n", idx, static_cast<int>(enable), static_cast<int>(saved));
        return saved ? ft_setsuccess : ft_fileerror;
    }
    case FIELD_DOCUMENT_PROTECTION:
    {
        if (!fieldValue) return ft_fieldempty;
        std::wstring encoded(static_cast<const wchar_t*>(fieldValue));

        std::string mode;
        std::wstring wPass;
        size_t sep = encoded.find(L'|');
        if (sep == std::wstring::npos) {
            mode = WideToUtf8(encoded);
        }
        else {
            mode = WideToUtf8(encoded.substr(0, sep));
            wPass = encoded.substr(sep + 1);
        }

        bool saved = UpdateSettingsXml(ansiPath, [&](tinyxml2::XMLDocument& doc, tinyxml2::XMLElement* root) {
            tinyxml2::XMLElement* prot = root->FirstChildElement("w:documentProtection");
            if (prot) root->DeleteChild(prot);

            if (mode == "No protection") return true;

            const char* editVal =
                mode == "Read-only" ? "readOnly" :
                mode == "Filling in forms" ? "forms" :
                mode == "Comments" ? "comments" :
                mode == "Tracked changes" ? "trackedChanges" : nullptr;
            if (!editVal) return false;

            prot = doc.NewElement("w:documentProtection");
            prot->SetAttribute("w:edit", editVal);
            prot->SetAttribute("w:enforcement", "1");

            if (!wPass.empty()) {
                std::string pass;
                for (wchar_t ch : wPass) pass += static_cast<char>(ch & 0xFF);
                int len = static_cast<int>(pass.size());
                WORD hash = 0;
                for (int i = len - 1; i >= 0; --i) {
                    hash ^= static_cast<WORD>(static_cast<unsigned char>(pass[i]));
                    for (int bit = 0; bit < 7; ++bit) {
                        if (hash & 0x4000) hash = static_cast<WORD>(((hash << 1) & 0x7FFF) ^ 0x6072);
                        else hash = static_cast<WORD>((hash << 1) & 0x7FFF);
                    }
                }
                hash ^= static_cast<WORD>(len);
                hash ^= 0xCE4B;

                char hashHex[8];
                snprintf(hashHex, sizeof(hashHex), "%04X", static_cast<unsigned>(hash));
                prot->SetAttribute("w:cryptAlgorithmClass", "hash");
                prot->SetAttribute("w:cryptAlgorithmType", "typeAny");
                prot->SetAttribute("w:cryptAlgorithmSid", "1");
                prot->SetAttribute("w:cryptSpinCount", "0");
                prot->SetAttribute("w:hash", hashHex);
                prot->SetAttribute("w:salt", "");
            }

            InsertSettingsElement(root, prot);
            return true;
        });

        return saved ? ft_setsuccess : ft_fileerror;
    }
    case FIELD_AUTHORS:
    {
        if (!fieldValue) return ft_fieldempty;
        std::wstring encoded(static_cast<const wchar_t*>(fieldValue));
        std::vector<AuthorRenameEntry> renames;
        std::string replaceRemainingWith;
        ParseAuthorRenamePayload(encoded, renames, replaceRemainingWith);

        bool hasUsefulSpecificRename = std::any_of(renames.begin(), renames.end(), [](const AuthorRenameEntry& rename) {
            return !rename.oldAuthor.empty() && !rename.newAuthor.empty();
        });

        if (!hasUsefulSpecificRename && replaceRemainingWith.empty()) return ft_fieldempty;
        return RenameTrackedChangeAuthorsBatch(ansiPath, renames, replaceRemainingWith) ? ft_setsuccess : ft_fileerror;
    }
    default:
        return ft_notsupported;
    }
}

int RunContentSetValue(char* fileName, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int flags)
{
    DbgLog("ContentSetValue called: fieldIndex=%d\n", fieldIndex);

    if (fieldIndex == -1) {
        ClearCache();
        return ft_setsuccess;
    }

    if (!fileName) return ft_fieldempty;

    std::wstring wfn;
    if (!AnsiToWideAcp(fileName, wfn)) return ft_fieldempty;

    const void* passValue = fieldValue;
    std::wstring widebuf;
    if (fieldValue && (fieldType == ft_string || fieldType == ft_stringw)) {
        if (AnsiToWideAcp(static_cast<const char*>(fieldValue), widebuf)) {
            passValue = widebuf.c_str();
        }
    }

    return RunContentSetValueW(&wfn[0], fieldIndex, unitIndex, fieldType, const_cast<void*>(passValue), flags);
}
