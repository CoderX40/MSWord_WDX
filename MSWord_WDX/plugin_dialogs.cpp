#include "pch.h"

#include "plugin_shared.h"

namespace {

struct AuthorDlgData {
    std::wstring oldAuthor;
    std::wstring newAuthor;
    bool accepted = false;
};

struct ProtDlgData {
    bool currentlyProtected = false;
    std::wstring chosenMode;
    std::wstring password;
    bool accepted = false;
};

#define ADL_OLD_EDIT 101
#define ADL_NEW_EDIT 102
#define PDL_COMBO   201
#define PDL_PASSLBL 202
#define PDL_PASS    203

INT_PTR CALLBACK AuthorDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
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

    Item(SS_LEFT, 7, 7, 246, 8, 0xFFFF, L"STATIC", L"Old Author Name (leave blank to rename ALL authors):");
    Item(ES_AUTOHSCROLL | WS_BORDER, 7, 18, 246, 14, ADL_OLD_EDIT, L"EDIT", L"");
    Item(SS_LEFT, 7, 40, 246, 8, 0xFFFF, L"STATIC", L"New Author Name:");
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
        SetWindowTextW(hDlg, L"Document Protection");
        HWND hCombo = GetDlgItem(hDlg, PDL_COMBO);
        SendMessageW(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"No protection"));
        if (!d->currentlyProtected) {
            SendMessageW(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Read-Only"));
            SendMessageW(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Forms"));
            SendMessageW(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Comments"));
            SendMessageW(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Tracked Changes"));
        }
        SendMessageW(hCombo, CB_SETCURSEL, 0, 0);
        ShowWindow(GetDlgItem(hDlg, PDL_PASSLBL), SW_HIDE);
        ShowWindow(GetDlgItem(hDlg, PDL_PASS), SW_HIDE);
        return TRUE;
    }
    case WM_COMMAND:
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
                    MessageBoxW(hDlg, L"Please enter a password.", L"Document Protection", MB_OK | MB_ICONWARNING);
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
    W(0); W(0); WS(L"Document Protection"); W(8); WS(L"MS Shell Dlg");

    auto Item = [&](DWORD style, short x, short y, short w, short h, WORD id, const wchar_t* cls, const wchar_t* text) {
        Al();
        DW(style | WS_CHILD | WS_VISIBLE); DW(0);
        W(static_cast<WORD>(x)); W(static_cast<WORD>(y)); W(static_cast<WORD>(w)); W(static_cast<WORD>(h)); W(id);
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

} // namespace

int RunContentEditValueW(HWND parentWin, int fieldIndex, int unitIndex, int fieldType, void* fieldValue, int maxLen, int flags, const char* langIdentifier)
{
    UNREFERENCED_PARAMETER(unitIndex);
    UNREFERENCED_PARAMETER(flags);
    UNREFERENCED_PARAMETER(langIdentifier);

    DbgLog("ContentEditValue called: fieldIndex=%d fieldType=%d\n", fieldIndex, fieldType);
    if (fieldIndex != FIELD_AUTHORS && fieldIndex != FIELD_DOCUMENT_PROTECTION) return ft_notsupported;
    if (!fieldValue) return ft_fieldempty;

    if (fieldIndex == FIELD_AUTHORS) {
        AuthorDlgData dlgData;
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
