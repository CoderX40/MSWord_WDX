#include "pch.h"

#include "plugin_shared.h"

#include <algorithm>
#include <cctype>
#include <cstdio>
#include <cstdlib>
#include <functional>
#include <map>
#include <vector>

namespace {

bool IsXmlWhitespace(char ch);
std::string EncodeXmlAttributeValue(const std::string& value, char quote);

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
    char suffix[64];
    snprintf(suffix, sizeof(suffix), ".%lu.%lu.tmp", static_cast<unsigned long>(GetCurrentProcessId()), static_cast<unsigned long>(GetTickCount()));
    std::string tmpPath = path + suffix;

    auto fail = [&]() {
        DeleteFileA(tmpPath.c_str());
        return false;
    };

    mz_zip_archive reader{};
    if (!mz_zip_reader_init_file(&reader, path.c_str(), 0)) return false;

    mz_zip_archive writer{};
    if (!mz_zip_writer_init_file(&writer, tmpPath.c_str(), 0)) {
        mz_zip_reader_end(&reader);
        return fail();
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
                return fail();
            }
        }
        else if (!mz_zip_writer_add_from_zip_reader(&writer, &reader, i)) {
            mz_zip_writer_end(&writer);
            mz_zip_reader_end(&reader);
            return fail();
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
                return fail();
            }
        }
    }

    if (!mz_zip_writer_finalize_archive(&writer)) {
        mz_zip_writer_end(&writer);
        mz_zip_reader_end(&reader);
        return fail();
    }

    mz_zip_writer_end(&writer);
    mz_zip_reader_end(&reader);

    auto validateEntry = [](mz_zip_archive& zip, const char* name) {
        return mz_zip_reader_locate_file(&zip, name, nullptr, 0) >= 0;
    };

    mz_zip_archive validator{};
    bool valid = mz_zip_reader_init_file(&validator, tmpPath.c_str(), 0) == MZ_TRUE;
    if (valid) {
        valid =
            validateEntry(validator, "[Content_Types].xml") &&
            validateEntry(validator, "_rels/.rels") &&
            validateEntry(validator, "word/document.xml");

        for (const auto& kv : replacements) {
            if (!validateEntry(validator, kv.first.c_str())) {
                valid = false;
                break;
            }
        }
    }

    if (valid) mz_zip_reader_end(&validator);
    if (!valid) {
        if (validator.m_zip_mode != MZ_ZIP_MODE_INVALID) mz_zip_reader_end(&validator);
        return fail();
    }

    if (MoveFileExA(tmpPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED) == FALSE) {
        return fail();
    }

    return true;
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

std::string GetXmlNameAt(const std::string& xml, size_t nameStart)
{
    size_t pos = nameStart;
    while (pos < xml.size() && !IsXmlWhitespace(xml[pos]) && xml[pos] != '/' && xml[pos] != '>') ++pos;
    return xml.substr(nameStart, pos - nameStart);
}

bool IsXmlNameEnd(char ch)
{
    return IsXmlWhitespace(ch) || ch == '/' || ch == '>';
}

size_t FindTagEnd(const std::string& xml, size_t tagStart)
{
    char quote = '\0';
    for (size_t i = tagStart; i < xml.size(); ++i) {
        char ch = xml[i];
        if (quote) {
            if (ch == quote) quote = '\0';
        }
        else if (ch == '"' || ch == '\'') {
            quote = ch;
        }
        else if (ch == '>') {
            return i;
        }
    }
    return std::string::npos;
}

bool IsSelfClosingTag(const std::string& xml, size_t tagStart, size_t tagEnd)
{
    size_t pos = tagEnd;
    while (pos > tagStart && IsXmlWhitespace(xml[pos - 1])) --pos;
    return pos > tagStart && xml[pos - 1] == '/';
}

bool FindElementRangeByNames(const std::string& xml, const std::vector<std::string>& names, size_t searchStart, size_t searchEnd, size_t& elementStart, size_t& elementEnd, std::string* matchedName = nullptr)
{
    size_t pos = searchStart;
    while ((pos = xml.find('<', pos)) != std::string::npos && pos < searchEnd) {
        if (pos + 1 >= searchEnd || xml[pos + 1] == '/' || xml[pos + 1] == '?' || xml[pos + 1] == '!') {
            ++pos;
            continue;
        }

        std::string name = GetXmlNameAt(xml, pos + 1);
        bool wanted = false;
        for (const std::string& candidate : names) {
            if (name == candidate) {
                wanted = true;
                break;
            }
        }

        if (!wanted) {
            ++pos;
            continue;
        }

        size_t tagEnd = FindTagEnd(xml, pos);
        if (tagEnd == std::string::npos || tagEnd >= searchEnd) return false;

        elementStart = pos;
        if (IsSelfClosingTag(xml, pos, tagEnd)) {
            elementEnd = tagEnd + 1;
        }
        else {
            std::string closingTag = "</" + name + ">";
            size_t closeStart = xml.find(closingTag, tagEnd + 1);
            if (closeStart == std::string::npos || closeStart >= searchEnd) return false;
            elementEnd = closeStart + closingTag.size();
        }

        if (matchedName) *matchedName = name;
        return true;
    }

    return false;
}

bool FindSettingsRoot(const std::string& xml, size_t& contentStart, size_t& contentEnd)
{
    size_t rootStart = 0;
    size_t rootEnd = 0;
    std::string rootName;
    if (!FindElementRangeByNames(xml, { "w:settings", "settings" }, 0, xml.size(), rootStart, rootEnd, &rootName)) return false;

    size_t openEnd = FindTagEnd(xml, rootStart);
    if (openEnd == std::string::npos || openEnd + 1 > rootEnd) return false;

    std::string closeTag = "</" + rootName + ">";
    if (rootEnd < closeTag.size()) return false;

    contentStart = openEnd + 1;
    contentEnd = rootEnd - closeTag.size();
    return contentStart <= contentEnd;
}

bool FindSettingsChild(const std::string& xml, const std::vector<std::string>& names, size_t& elementStart, size_t& elementEnd)
{
    size_t contentStart = 0;
    size_t contentEnd = 0;
    if (!FindSettingsRoot(xml, contentStart, contentEnd)) return false;
    return FindElementRangeByNames(xml, names, contentStart, contentEnd, elementStart, elementEnd);
}

size_t FindSettingsInsertPosition(const std::string& xml, const char* newElementName)
{
    size_t contentStart = 0;
    size_t contentEnd = 0;
    if (!FindSettingsRoot(xml, contentStart, contentEnd)) return std::string::npos;

    int newRank = GetSettingsElementPriority(newElementName);
    size_t pos = contentStart;
    while ((pos = xml.find('<', pos)) != std::string::npos && pos < contentEnd) {
        if (pos + 1 >= contentEnd || xml[pos + 1] == '/' || xml[pos + 1] == '?' || xml[pos + 1] == '!') {
            ++pos;
            continue;
        }

        std::string name = GetXmlNameAt(xml, pos + 1);
        if (GetSettingsElementPriority(name.c_str()) > newRank) return pos;

        size_t childStart = 0;
        size_t childEnd = 0;
        if (!FindElementRangeByNames(xml, { name }, pos, contentEnd, childStart, childEnd)) return std::string::npos;
        pos = childEnd;
    }

    return contentEnd;
}

bool SaveSettingsXml(const std::string& ansiPath, const std::string& settingsXml)
{
    bool saved = SaveXmlToZip(ansiPath.c_str(), "word/settings.xml", settingsXml);
    if (saved) ClearCache();
    return saved;
}

bool SetSettingsElementXml(std::string& settingsXml, const std::vector<std::string>& names, const char* insertName, const std::string& elementXml, bool enable)
{
    size_t elementStart = 0;
    size_t elementEnd = 0;
    if (FindSettingsChild(settingsXml, names, elementStart, elementEnd)) {
        if (enable) {
            settingsXml.replace(elementStart, elementEnd - elementStart, elementXml);
        }
        else {
            settingsXml.erase(elementStart, elementEnd - elementStart);
        }
        return true;
    }

    if (!enable) return true;

    size_t insertPos = FindSettingsInsertPosition(settingsXml, insertName);
    if (insertPos == std::string::npos) return false;
    settingsXml.insert(insertPos, elementXml);
    return true;
}

bool UpdateSettingsOnOffElement(const std::string& ansiPath, const char* localName, bool enable)
{
    std::string settingsXml;
    if (!ExtractRawFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml)) return false;

    std::string fullName = std::string("w:") + localName;
    std::string elementXml = "<" + fullName + " w:val=\"true\"/>";
    if (!SetSettingsElementXml(settingsXml, { fullName, localName }, fullName.c_str(), elementXml, enable)) return false;

    return SaveSettingsXml(ansiPath, settingsXml);
}

bool UpdateAnonymisationElements(const std::string& ansiPath, bool removePI, bool removeDate)
{
    std::string settingsXml;
    if (!ExtractRawFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml)) return false;

    if (!SetSettingsElementXml(settingsXml,
        { "w:removePersonalInformation", "removePersonalInformation" },
        "w:removePersonalInformation",
        "<w:removePersonalInformation w:val=\"true\"/>",
        removePI)) {
        return false;
    }

    if (!SetSettingsElementXml(settingsXml,
        { "w:removeDateAndTime", "removeDateAndTime" },
        "w:removeDateAndTime",
        "<w:removeDateAndTime w:val=\"true\"/>",
        removeDate)) {
        return false;
    }

    return SaveSettingsXml(ansiPath, settingsXml);
}

std::string BuildDocumentProtectionXml(const std::string& editVal, const std::string& hashHex)
{
    std::string xml = "<w:documentProtection w:edit=\"" + EncodeXmlAttributeValue(editVal, '"') + "\" w:enforcement=\"1\"";
    if (!hashHex.empty()) {
        xml += " w:cryptAlgorithmClass=\"hash\" w:cryptAlgorithmType=\"typeAny\" w:cryptAlgorithmSid=\"1\" w:cryptSpinCount=\"0\" w:hash=\"";
        xml += EncodeXmlAttributeValue(hashHex, '"');
        xml += "\" w:salt=\"\"";
    }
    xml += "/>";
    return xml;
}

bool UpdateDocumentProtectionElement(const std::string& ansiPath, const std::string& mode, const std::string& hashHex)
{
    std::string settingsXml;
    if (!ExtractRawFileFromZip(ansiPath.c_str(), "word/settings.xml", settingsXml)) return false;

    if (mode == "No protection") {
        if (!SetSettingsElementXml(settingsXml, { "w:documentProtection", "documentProtection" }, "w:documentProtection", std::string(), false)) return false;
        return SaveSettingsXml(ansiPath, settingsXml);
    }

    const char* editVal =
        mode == "Read-only" ? "readOnly" :
        mode == "Filling in forms" ? "forms" :
        mode == "Comments" ? "comments" :
        mode == "Tracked changes" ? "trackedChanges" : nullptr;
    if (!editVal) return false;

    std::string elementXml = BuildDocumentProtectionXml(editVal, hashHex);
    if (!SetSettingsElementXml(settingsXml, { "w:documentProtection", "documentProtection" }, "w:documentProtection", elementXml, true)) return false;

    return SaveSettingsXml(ansiPath, settingsXml);
}

std::string EncodeXmlTextValue(const std::string& value)
{
    std::string encoded;
    encoded.reserve(value.size());

    for (char ch : value) {
        switch (ch) {
        case '&': encoded += "&amp;"; break;
        case '<': encoded += "&lt;"; break;
        case '>': encoded += "&gt;"; break;
        default: encoded.push_back(ch); break;
        }
    }

    return encoded;
}

bool FindRootContentByNames(const std::string& xml, const std::vector<std::string>& rootNames, size_t& rootStart, size_t& openEnd, size_t& contentStart, size_t& contentEnd)
{
    size_t rootEnd = 0;
    std::string rootName;
    if (!FindElementRangeByNames(xml, rootNames, 0, xml.size(), rootStart, rootEnd, &rootName)) return false;

    openEnd = FindTagEnd(xml, rootStart);
    if (openEnd == std::string::npos || openEnd + 1 > rootEnd) return false;

    std::string closeTag = "</" + rootName + ">";
    if (rootEnd < closeTag.size()) return false;

    contentStart = openEnd + 1;
    contentEnd = rootEnd - closeTag.size();
    return contentStart <= contentEnd;
}

bool EnsureRootNamespace(std::string& xml, size_t rootStart, size_t& openEnd, const char* prefix, const char* uri)
{
    if (!prefix || !*prefix || !uri || !*uri) return true;

    std::string attrName = std::string("xmlns:") + prefix;
    size_t nameStart = rootStart + 1;
    std::string rootName = GetXmlNameAt(xml, nameStart);
    size_t attrSearchStart = nameStart + rootName.size();

    size_t pos = attrSearchStart;
    while ((pos = xml.find(attrName, pos)) != std::string::npos && pos < openEnd) {
        bool leftOk = pos == attrSearchStart || IsXmlWhitespace(xml[pos - 1]);
        size_t right = pos + attrName.size();
        bool rightOk = right < openEnd && (IsXmlWhitespace(xml[right]) || xml[right] == '=');
        if (leftOk && rightOk) return true;
        pos = right;
    }

    std::string attr = " " + attrName + "=\"" + EncodeXmlAttributeValue(uri, '"') + "\"";
    xml.insert(nameStart + rootName.size(), attr);
    openEnd += attr.size();
    return true;
}

bool EnsureNamespaceForElement(std::string& xml, size_t rootStart, size_t& openEnd, const char* elementName)
{
    const char* colon = strchr(elementName, ':');
    if (!colon) return true;

    std::string prefix(elementName, colon - elementName);
    if (prefix == "dc") return EnsureRootNamespace(xml, rootStart, openEnd, "dc", "http://purl.org/dc/elements/1.1/");
    if (prefix == "cp") return EnsureRootNamespace(xml, rootStart, openEnd, "cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    if (prefix == "dcterms") return EnsureRootNamespace(xml, rootStart, openEnd, "dcterms", "http://purl.org/dc/terms/");
    return true;
}

bool SetDirectChildElementText(std::string& xml, const std::vector<std::string>& rootNames, const char* elementName, const std::string& value)
{
    size_t rootStart = 0;
    size_t openEnd = 0;
    size_t contentStart = 0;
    size_t contentEnd = 0;
    if (!FindRootContentByNames(xml, rootNames, rootStart, openEnd, contentStart, contentEnd)) return false;

    std::string encoded = EncodeXmlTextValue(value);
    size_t elementStart = 0;
    size_t elementEnd = 0;
    if (FindElementRangeByNames(xml, { elementName }, contentStart, contentEnd, elementStart, elementEnd)) {
        size_t elementOpenEnd = FindTagEnd(xml, elementStart);
        if (elementOpenEnd == std::string::npos || elementOpenEnd >= elementEnd) return false;

        if (IsSelfClosingTag(xml, elementStart, elementOpenEnd)) {
            size_t slash = elementOpenEnd;
            while (slash > elementStart && IsXmlWhitespace(xml[slash - 1])) --slash;
            if (slash <= elementStart || xml[slash - 1] != '/') return false;
            size_t slashIndex = slash - 1;
            xml.replace(slashIndex, elementOpenEnd - slashIndex + 1, std::string(">") + encoded + "</" + elementName + ">");
        }
        else {
            std::string closeTag = "</" + std::string(elementName) + ">";
            if (elementEnd < closeTag.size()) return false;
            size_t valueStart = elementOpenEnd + 1;
            size_t valueEnd = elementEnd - closeTag.size();
            xml.replace(valueStart, valueEnd - valueStart, encoded);
        }
        return true;
    }

    if (!EnsureNamespaceForElement(xml, rootStart, openEnd, elementName)) return false;

    if (!FindRootContentByNames(xml, rootNames, rootStart, openEnd, contentStart, contentEnd)) return false;
    std::string elementXml = "<" + std::string(elementName) + ">" + encoded + "</" + std::string(elementName) + ">";
    xml.insert(contentEnd, elementXml);
    return true;
}

bool SetPackageXmlStringValue(std::string& xmlContent, const char* fileNameInZip, const char* elementName, const std::string& value)
{
    if (!fileNameInZip || !elementName) return false;

    if (strcmp(fileNameInZip, "docProps/core.xml") == 0) {
        return SetDirectChildElementText(xmlContent, { "cp:coreProperties", "coreProperties" }, elementName, value);
    }

    if (strcmp(fileNameInZip, "docProps/app.xml") == 0) {
        return SetDirectChildElementText(xmlContent, { "Properties" }, elementName, value);
    }

    return SetXmlStringValue(xmlContent, elementName, value);
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

bool IsXmlWhitespace(char ch)
{
    return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n';
}

void AppendUtf8CodePoint(std::string& output, unsigned long codePoint)
{
    if (codePoint <= 0x7F) {
        output.push_back(static_cast<char>(codePoint));
    }
    else if (codePoint <= 0x7FF) {
        output.push_back(static_cast<char>(0xC0 | (codePoint >> 6)));
        output.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
    }
    else if (codePoint <= 0xFFFF) {
        output.push_back(static_cast<char>(0xE0 | (codePoint >> 12)));
        output.push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
        output.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
    }
    else if (codePoint <= 0x10FFFF) {
        output.push_back(static_cast<char>(0xF0 | (codePoint >> 18)));
        output.push_back(static_cast<char>(0x80 | ((codePoint >> 12) & 0x3F)));
        output.push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
        output.push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
    }
}

std::string DecodeXmlAttributeValue(const std::string& value)
{
    std::string decoded;
    decoded.reserve(value.size());

    for (size_t i = 0; i < value.size(); ++i) {
        if (value[i] != '&') {
            decoded.push_back(value[i]);
            continue;
        }

        size_t semi = value.find(';', i + 1);
        if (semi == std::string::npos) {
            decoded.push_back(value[i]);
            continue;
        }

        std::string entity = value.substr(i + 1, semi - i - 1);
        if (entity == "amp") decoded.push_back('&');
        else if (entity == "lt") decoded.push_back('<');
        else if (entity == "gt") decoded.push_back('>');
        else if (entity == "quot") decoded.push_back('"');
        else if (entity == "apos") decoded.push_back('\'');
        else if (!entity.empty() && entity[0] == '#') {
            char* end = nullptr;
            unsigned long codePoint = 0;
            if (entity.size() > 2 && (entity[1] == 'x' || entity[1] == 'X')) {
                codePoint = strtoul(entity.c_str() + 2, &end, 16);
            }
            else {
                codePoint = strtoul(entity.c_str() + 1, &end, 10);
            }

            if (end && *end == '\0') AppendUtf8CodePoint(decoded, codePoint);
            else decoded.append(value, i, semi - i + 1);
        }
        else {
            decoded.append(value, i, semi - i + 1);
        }

        i = semi;
    }

    return decoded;
}

std::string EncodeXmlAttributeValue(const std::string& value, char quote)
{
    std::string encoded;
    encoded.reserve(value.size());

    for (char ch : value) {
        switch (ch) {
        case '&': encoded += "&amp;"; break;
        case '<': encoded += "&lt;"; break;
        case '>': encoded += "&gt;"; break;
        case '"':
            if (quote == '"') encoded += "&quot;";
            else encoded.push_back(ch);
            break;
        case '\'':
            if (quote == '\'') encoded += "&apos;";
            else encoded.push_back(ch);
            break;
        default:
            encoded.push_back(ch);
            break;
        }
    }

    return encoded;
}

const std::string* FindAuthorReplacement(const std::string& author, const std::vector<AuthorRenameEntry>& renames, const std::string& replaceRemainingWith)
{
    for (const AuthorRenameEntry& rename : renames) {
        if (rename.oldAuthor == author) return &rename.newAuthor;
    }

    return replaceRemainingWith.empty() ? nullptr : &replaceRemainingWith;
}

bool ReplaceAuthorAttributeInXml(std::string& xml, const char* attrName, const std::vector<AuthorRenameEntry>& renames, const std::string& replaceRemainingWith)
{
    bool modified = false;
    size_t pos = 0;
    const size_t attrLen = strlen(attrName);

    while ((pos = xml.find(attrName, pos)) != std::string::npos) {
        if (pos > 0 && !IsXmlWhitespace(xml[pos - 1])) {
            pos += attrLen;
            continue;
        }

        size_t cursor = pos + attrLen;
        while (cursor < xml.size() && IsXmlWhitespace(xml[cursor])) ++cursor;
        if (cursor >= xml.size() || xml[cursor] != '=') {
            pos += attrLen;
            continue;
        }

        ++cursor;
        while (cursor < xml.size() && IsXmlWhitespace(xml[cursor])) ++cursor;
        if (cursor >= xml.size() || (xml[cursor] != '"' && xml[cursor] != '\'')) {
            pos += attrLen;
            continue;
        }

        char quote = xml[cursor];
        size_t valueStart = cursor + 1;
        size_t valueEnd = xml.find(quote, valueStart);
        if (valueEnd == std::string::npos) break;

        std::string rawValue = xml.substr(valueStart, valueEnd - valueStart);
        std::string decodedValue = DecodeXmlAttributeValue(rawValue);
        const std::string* replacement = FindAuthorReplacement(decodedValue, renames, replaceRemainingWith);

        if (replacement && *replacement != decodedValue) {
            std::string encodedReplacement = EncodeXmlAttributeValue(*replacement, quote);
            xml.replace(valueStart, valueEnd - valueStart, encodedReplacement);
            modified = true;
            pos = valueStart + encodedReplacement.size() + 1;
        }
        else {
            pos = valueEnd + 1;
        }
    }

    return modified;
}

bool ReplaceTrackedChangeAuthorAttributesInXml(std::string& xml, const std::vector<AuthorRenameEntry>& renames, const std::string& replaceRemainingWith)
{
    bool modified = ReplaceAuthorAttributeInXml(xml, "w:author", renames, replaceRemainingWith);
    modified = ReplaceAuthorAttributeInXml(xml, "w:originalAuthor", renames, replaceRemainingWith) || modified;
    return modified;
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

        if (ReplaceTrackedChangeAuthorAttributesInXml(xml, renames, hasReplaceRemaining ? replaceRemainingWith : std::string())) {
            replacements[stat.m_filename] = xml;
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
            if (!ExtractRawFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) return ft_fileerror;

            const char* elemName =
                fieldIndex == FIELD_CORE_TITLE ? "dc:title" :
                fieldIndex == FIELD_CORE_SUBJECT ? "dc:subject" :
                fieldIndex == FIELD_CORE_CREATOR ? "dc:creator" :
                fieldIndex == FIELD_CORE_KEYWORDS ? "cp:keywords" :
                "dc:description";

            if (!SetPackageXmlStringValue(coreXml, "docProps/core.xml", elemName, value) || !SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
                return ft_fileerror;
            }
            ClearCache();
        }
        else {
            std::string appXml;
            if (!ExtractRawFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml)) return ft_fileerror;

            const char* elemName =
                fieldIndex == FIELD_APP_MANAGER ? "Manager" :
                fieldIndex == FIELD_APP_COMPANY ? "Company" :
                "HyperlinkBase";

            if (!SetPackageXmlStringValue(appXml, "docProps/app.xml", elemName, value) || !SaveXmlToZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
                return ft_fileerror;
            }
            ClearCache();
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
        if (!ExtractRawFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) return ft_fileerror;
        const char* elemName =
            fieldIndex == FIELD_CORE_CREATED_DATE ? "dcterms:created" :
            fieldIndex == FIELD_CORE_MODIFIED_DATE ? "dcterms:modified" :
            "cp:lastPrinted";

        if (!SetPackageXmlStringValue(coreXml, "docProps/core.xml", elemName, newValue) || !SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            return ft_fileerror;
        }
        ClearCache();
        return ft_setsuccess;
    }
    case FIELD_CORE_LAST_MODIFIED_BY:
    {
        std::string userName = FieldValueToUtf8(fieldType, fieldValue);
        std::string coreXml;
        if (!ExtractRawFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) return ft_fileerror;
        if (!SetPackageXmlStringValue(coreXml, "docProps/core.xml", "cp:lastModifiedBy", userName) || !SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            return ft_fileerror;
        }
        ClearCache();
        return ft_setsuccess;
    }
    case FIELD_CORE_REVISION_NUMBER:
    {
        if (!fieldValue) return ft_fieldempty;
        int revision = *static_cast<const int*>(fieldValue);
        std::string coreXml;
        if (!ExtractRawFileFromZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) return ft_fileerror;
        char buf[16];
        snprintf(buf, sizeof(buf), "%d", revision);
        if (!SetPackageXmlStringValue(coreXml, "docProps/core.xml", "cp:revision", buf) || !SaveXmlToZip(ansiPath.c_str(), "docProps/core.xml", coreXml)) {
            return ft_fileerror;
        }
        ClearCache();
        return ft_setsuccess;
    }
    case FIELD_APP_EDITING_TIME:
    {
        if (!fieldValue) return ft_fieldempty;
        int editingTime = *static_cast<const int*>(fieldValue);
        std::string appXml;
        if (!ExtractRawFileFromZip(ansiPath.c_str(), "docProps/app.xml", appXml)) return ft_fileerror;
        char buf[16];
        snprintf(buf, sizeof(buf), "%d", editingTime);
        if (!SetPackageXmlStringValue(appXml, "docProps/app.xml", "TotalTime", buf) || !SaveXmlToZip(ansiPath.c_str(), "docProps/app.xml", appXml)) {
            return ft_fileerror;
        }
        ClearCache();
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

        bool saved = UpdateSettingsOnOffElement(ansiPath, "linkStyles", enable);

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

        bool saved = UpdateAnonymisationElements(ansiPath, removePI, removeDate);

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

        bool saved = UpdateSettingsOnOffElement(ansiPath, "trackRevisions", enable);

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

        std::string hashHex;
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

            char hashBuf[8];
            snprintf(hashBuf, sizeof(hashBuf), "%04X", static_cast<unsigned>(hash));
            hashHex = hashBuf;
        }

        bool saved = UpdateDocumentProtectionElement(ansiPath, mode, hashHex);

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
