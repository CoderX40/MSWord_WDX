#include "pch.h"

#include "plugin_shared.h"

namespace {

struct AuthorDlgData {
    std::wstring oldAuthor;
    std::wstring newAuthor;
    std::string langIdentifier;
    bool accepted = false;
};

struct ProtDlgData {
    bool currentlyProtected = false;
    std::wstring chosenMode;
    std::wstring password;
    std::string langIdentifier;
    bool accepted = false;
};

#define ADL_OLD_EDIT 101
#define ADL_NEW_EDIT 102
#define PDL_COMBO   201
#define PDL_PASSLBL 202
#define PDL_PASS    203
#define ADL_OLD_LABEL 301
#define ADL_NEW_LABEL 302
#define PDL_MODE_LABEL 303

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
    if (dot == std::wstring::npos) {
        result += L".lng";
    }
    else {
        result.erase(dot);
        result += L".lng";
    }

    return result;
}

std::wstring WidenAscii(const std::string& text)
{
    return std::wstring(text.begin(), text.end());
}

std::wstring Utf8ToWideBestEffort(const std::string& text)
{
    std::wstring wide;
    if (Utf8ToWideString(text, wide)) {
        return wide;
    }
    return WidenAscii(text);
}

std::wstring TranslateDialogText(const std::string& langIdentifier, const wchar_t* fallback)
{
    if (!fallback || !*fallback || langIdentifier.empty()) {
        return fallback ? std::wstring(fallback) : std::wstring();
    }

    std::wstring lngPath = GetPluginLngPath();
    if (lngPath.empty()) {
        return std::wstring(fallback);
    }

    std::string keyUtf8 = WideToUtf8(std::wstring(fallback));
    if (keyUtf8.empty()) {
        return std::wstring(fallback);
    }

    std::wstring section = WidenAscii(langIdentifier);
    std::wstring key = Utf8ToWideBestEffort(keyUtf8);
    wchar_t buffer[1024] = {};
    GetPrivateProfileStringW(section.c_str(), key.c_str(), fallback, buffer, static_cast<DWORD>(std::size(buffer)), lngPath.c_str());
    return std::wstring(buffer);
}

void SetDlgItemTextLng(HWND hDlg, int controlId, const std::string& langIdentifier, const wchar_t* fallback)
{
    SetWindowTextW(GetDlgItem(hDlg, controlId), TranslateDialogText(langIdentifier, fallback).c_str());
}

int AddComboTextLng(HWND combo, const std::string& langIdentifier, const wchar_t* fallback)
{
    std::wstring text = TranslateDialogText(langIdentifier, fallback);
    return static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str())));
}

INT_PTR CALLBACK AuthorDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_INITDIALOG:
    {
        SetWindowLongPtr(hDlg, GWLP_USERDATA, lParam);
        AuthorDlgData* d = reinterpret_cast<AuthorDlgData*>(lParam);
        SetWindowTextW(hDlg, TranslateDialogText(d->langIdentifier, L"Change tracked change author name").c_str());
        SetDlgItemTextLng(hDlg, ADL_OLD_LABEL, d->langIdentifier, L"Old author name (leave blank to rename all):");
        SetDlgItemTextLng(hDlg, ADL_NEW_LABEL, d->langIdentifier, L"New author name:");
        SetDlgItemTextLng(hDlg, IDOK, d->langIdentifier, L"OK");
        SetDlgItemTextLng(hDlg, IDCANCEL, d->langIdentifier, L"Cancel");
        return TRUE;
    }
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
                MessageBoxW(
                    hDlg,
                    TranslateDialogText(d->langIdentifier, L"Please enter a new name.").c_str(),
                    TranslateDialogText(d->langIdentifier, L"Error").c_str(),
                    MB_OK | MB_ICONWARNING);
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

DLGTEMPLATE* BuildAuthorDlgTemplate()
{
    static WORD buf[2048];
    memset(buf, 0, sizeof(buf));
    WORD* p = buf;

    auto W = [&](WORD w) { *p++ = w; };
    auto DW = [&](DWORD dw) { memcpy(p, &dw, 4); p += 2; };
    auto WS = [&](const wchar_t* s) { while (*s) *p++ = static_cast<WORD>(*s++); *p++ = 0; };
    auto Al = [&]() { if ((DWORD_PTR)p & 2) *p++ = 0; };

    DW(DS_SETFONT | DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU);
    DW(0); W(6); W(0); W(0); W(260); W(120);
    W(0); W(0); WS(L"Rename Author"); W(8); WS(L"MS Shell Dlg");

    auto Item = [&](DWORD style, short x, short y, short w, short h, WORD id, const wchar_t* cls, const wchar_t* text) {
        Al();
        DW(style | WS_CHILD | WS_VISIBLE); DW(0);
        W(static_cast<WORD>(x)); W(static_cast<WORD>(y)); W(static_cast<WORD>(w)); W(static_cast<WORD>(h)); W(id);
        WS(cls); WS(text); W(0);
    };

    Item(SS_LEFT, 7, 7, 246, 8, ADL_OLD_LABEL, L"STATIC", L"Old author name (leave blank to rename all):");
    Item(ES_AUTOHSCROLL | WS_BORDER, 7, 18, 246, 14, ADL_OLD_EDIT, L"EDIT", L"");
    Item(SS_LEFT, 7, 40, 246, 8, ADL_NEW_LABEL, L"STATIC", L"New author name:");
    Item(ES_AUTOHSCROLL | WS_BORDER, 7, 51, 246, 14, ADL_NEW_EDIT, L"EDIT", L"");
    Item(BS_DEFPUSHBUTTON, 148, 80, 50, 14, IDOK, L"BUTTON", L"OK");
    Item(BS_PUSHBUTTON, 203, 80, 50, 14, IDCANCEL, L"BUTTON", L"Cancel");

    return reinterpret_cast<DLGTEMPLATE*>(buf);
}

INT_PTR CALLBACK ProtDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_INITDIALOG:
    {
        SetWindowLongPtr(hDlg, GWLP_USERDATA, lParam);
        ProtDlgData* d = reinterpret_cast<ProtDlgData*>(lParam);
        SetWindowTextW(hDlg, TranslateDialogText(d->langIdentifier, L"Document protection").c_str());
        HWND hCombo = GetDlgItem(hDlg, PDL_COMBO);
        AddComboTextLng(hCombo, d->langIdentifier, L"No protection");
        if (!d->currentlyProtected) {
            AddComboTextLng(hCombo, d->langIdentifier, L"Read-only");
            AddComboTextLng(hCombo, d->langIdentifier, L"Filling in forms");
            AddComboTextLng(hCombo, d->langIdentifier, L"Comments");
            AddComboTextLng(hCombo, d->langIdentifier, L"Tracked changes");
        }
        SendMessageW(hCombo, CB_SETCURSEL, 0, 0);
        ShowWindow(GetDlgItem(hDlg, PDL_PASSLBL), SW_HIDE);
        ShowWindow(GetDlgItem(hDlg, PDL_PASS), SW_HIDE);
        SetDlgItemTextLng(hDlg, PDL_MODE_LABEL, d->langIdentifier, L"Restriction type:");
        SetDlgItemTextLng(hDlg, PDL_PASSLBL, d->langIdentifier, L"Password:");
        SetDlgItemTextLng(hDlg, IDOK, d->langIdentifier, L"OK");
        SetDlgItemTextLng(hDlg, IDCANCEL, d->langIdentifier, L"Cancel");
        return TRUE;
    }
    case WM_COMMAND:
    {
        if (HIWORD(wParam) == CBN_SELCHANGE && LOWORD(wParam) == PDL_COMBO) {
            int sel = static_cast<int>(SendMessageW(GetDlgItem(hDlg, PDL_COMBO), CB_GETCURSEL, 0, 0));
            bool needPass = sel > 0;
            ShowWindow(GetDlgItem(hDlg, PDL_PASSLBL), needPass ? SW_SHOW : SW_HIDE);
            ShowWindow(GetDlgItem(hDlg, PDL_PASS), needPass ? SW_SHOW : SW_HIDE);
            return TRUE;
        }
        if (LOWORD(wParam) == IDOK) {
            ProtDlgData* d = reinterpret_cast<ProtDlgData*>(GetWindowLongPtr(hDlg, GWLP_USERDATA));
            HWND hCombo = GetDlgItem(hDlg, PDL_COMBO);
            int sel = static_cast<int>(SendMessageW(hCombo, CB_GETCURSEL, 0, 0));
            wchar_t modeBuf[64] = {};
            SendMessageW(hCombo, CB_GETLBTEXT, sel, reinterpret_cast<LPARAM>(modeBuf));
            d->chosenMode = modeBuf;
            if (sel > 0) {
                wchar_t passBuf[256] = {};
                GetWindowTextW(GetDlgItem(hDlg, PDL_PASS), passBuf, 256);
                if (!passBuf[0]) {
                    MessageBoxW(
                        hDlg,
                        TranslateDialogText(d->langIdentifier, L"Please enter a password.").c_str(),
                        TranslateDialogText(d->langIdentifier, L"Document protection").c_str(),
                        MB_OK | MB_ICONWARNING);
                    return TRUE;
                }
                d->password = passBuf;
            }
            d->accepted = true;
            EndDialog(hDlg, IDOK);
            return TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        break;
    }
    case WM_CLOSE:
        EndDialog(hDlg, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

DLGTEMPLATE* BuildProtDlgTemplate()
{
    static WORD buf[2048];
    memset(buf, 0, sizeof(buf));
    WORD* p = buf;
    auto W = [&](WORD w) { *p++ = w; };
    auto DW = [&](DWORD dw) { memcpy(p, &dw, 4); p += 2; };
    auto WS = [&](const wchar_t* s) { while (*s) *p++ = static_cast<WORD>(*s++); *p++ = 0; };
    auto Al = [&]() { if ((DWORD_PTR)p & 2) *p++ = 0; };

    DW(DS_SETFONT | DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU);
    DW(0); W(6); W(0); W(0); W(240); W(105);
    W(0); W(0); WS(L"Document protection"); W(8); WS(L"MS Shell Dlg");

    auto Item = [&](DWORD style, short x, short y, short w, short h, WORD id, const wchar_t* cls, const wchar_t* text) {
        Al();
        DW(style | WS_CHILD | WS_VISIBLE); DW(0);
        W(static_cast<WORD>(x)); W(static_cast<WORD>(y)); W(static_cast<WORD>(w)); W(static_cast<WORD>(h)); W(id);
        WS(cls); WS(text); W(0);
    };

    Item(SS_LEFT, 7, 7, 226, 8, PDL_MODE_LABEL, L"STATIC", L"Restriction type:");
    Item(CBS_DROPDOWNLIST | WS_VSCROLL, 7, 18, 226, 80, PDL_COMBO, L"COMBOBOX", L"");
    Item(SS_LEFT, 7, 50, 226, 8, PDL_PASSLBL, L"STATIC", L"Password:");
    Item(ES_AUTOHSCROLL | ES_PASSWORD | WS_BORDER, 7, 61, 226, 14, PDL_PASS, L"EDIT", L"");
    Item(BS_DEFPUSHBUTTON, 130, 83, 50, 14, IDOK, L"BUTTON", L"OK");
    Item(BS_PUSHBUTTON, 183, 83, 50, 14, IDCANCEL, L"BUTTON", L"Cancel");

    return reinterpret_cast<DLGTEMPLATE*>(buf);
}

} // namespace

int RunContentEditValueW(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int maxLen, int flags, const char* langIdentifier)
{
    UNREFERENCED_PARAMETER(unitIndex);
    UNREFERENCED_PARAMETER(flags);

    DbgLog("ContentEditValue called: fieldIndex=%d fieldType=%d\n", fieldIndex, fieldType);
    if (fieldIndex != FIELD_AUTHORS && fieldIndex != FIELD_DOCUMENT_PROTECTION) return ft_notsupported;
    if (!fieldValue) return ft_fieldempty;

    if (fieldIndex == FIELD_AUTHORS) {
        AuthorDlgData dlgData;
        dlgData.langIdentifier = langIdentifier ? langIdentifier : "";
        if (DialogBoxIndirectParamW(GetModuleHandleW(nullptr), BuildAuthorDlgTemplate(), parentWin, AuthorDlgProc, reinterpret_cast<LPARAM>(&dlgData)) != IDOK || !dlgData.accepted) {
            return ft_setcancel;
        }

        std::wstring encoded = dlgData.oldAuthor + L"|" + dlgData.newAuthor;
        if (fieldType == ft_stringw) {
            wcsncpy_s(static_cast<WCHAR*>(fieldValue), maxLen / sizeof(WCHAR), encoded.c_str(), _TRUNCATE);
        }
        return ft_setsuccess;
    }

    ProtDlgData dlgData;
    dlgData.currentlyProtected = false;
    dlgData.langIdentifier = langIdentifier ? langIdentifier : "";

    if (DialogBoxIndirectParamW(GetModuleHandleW(nullptr), BuildProtDlgTemplate(), parentWin, ProtDlgProc, reinterpret_cast<LPARAM>(&dlgData)) != IDOK || !dlgData.accepted) {
        return ft_setcancel;
    }

    std::wstring encoded = dlgData.chosenMode + L"|" + dlgData.password;
    if (fieldType == ft_stringw) {
        wcsncpy_s(static_cast<WCHAR*>(fieldValue), maxLen / sizeof(WCHAR), encoded.c_str(), _TRUNCATE);
    }
    return ft_setsuccess;
}

int RunContentEditValue(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, char* fieldValue, int maxLen, int flags, const char* langIdentifier)
{
    DbgLog("ContentEditValue called: fieldIndex=%d fieldType=%d\n", fieldIndex, fieldType);
    if (fieldIndex != FIELD_AUTHORS && fieldIndex != FIELD_DOCUMENT_PROTECTION) return ft_notsupported;

    return RunContentEditValueW(parentWin, fieldIndex, unitIndex, fieldType, static_cast<void*>(fieldValue), maxLen, flags, langIdentifier);
}
