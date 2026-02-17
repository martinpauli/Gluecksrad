/*
 * Glücksrad – Faire Zufallsauswahl  (C++ / Win32 Native)
 * =======================================================
 * A faithful port of the Python tkinter "Glücksrad" application.
 *
 * Features:
 *   - Load names + counters from CSV (UTF-8, semicolon or comma separated)
 *   - Fair random selection (lowest-counter candidates only)
 *   - Animated "spin" through the list with deceleration
 *   - Winner blink highlight
 *   - Multi-draw with batch duplicate avoidance
 *   - Auto-save counters back to the CSV
 *   - Configurable spin parameters via Settings dialog
 *
 * File format:  CSV with header row.  Column 1 = Name, Column 2 = Counter.
 *   If column 2 is missing, counters default to 0.
 *   Semicolon (;) and comma (,) are both accepted as delimiters.
 *
 * Build (MSVC):
 *   cl /O2 /EHsc /DUNICODE /D_UNICODE gluecksrad.cpp /Fe:gluecksrad.exe
 *      user32.lib gdi32.lib comctl32.lib comdlg32.lib shell32.lib
 *
 * Build (MinGW-w64):
 *   g++ -O2 -mwindows -DUNICODE -D_UNICODE gluecksrad.cpp
 *      -o gluecksrad.exe -lcomctl32 -lcomdlg32 -lgdi32
 *
 * Build (CMake): see accompanying CMakeLists.txt
 */

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>

#include <algorithm>
#include <cstdlib>
#include <ctime>
#include <fstream>
#include <random>
#include <set>
#include <sstream>
#include <string>
#include <vector>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")

// ============================================================
//  Forward declarations
// ============================================================
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

// ============================================================
//  Data model
// ============================================================
struct NameEntry {
    std::wstring name;
    int counter = 0;
};

// ============================================================
//  Global application state
// ============================================================
static HINSTANCE g_hInst = nullptr;
static HWND g_hWnd = nullptr;         // main window
static HWND g_hList = nullptr;        // ListView
static HWND g_hBtnLoad = nullptr;
static HWND g_hBtnDraw = nullptr;
static HWND g_hBtnClear = nullptr;
static HWND g_hBtnReload = nullptr;
static HWND g_hEditN = nullptr;       // Spin control for draw count

static HWND g_hStatus = nullptr;      // Status bar

static std::vector<NameEntry> g_entries;
static std::wstring g_filePath;

// Animation state
static bool g_animRunning = false;
static int g_toDrawTotal = 0;
static int g_drawnCount = 0;
static std::set<int> g_roundSelectedIdx;
static std::set<int> g_roundExcludedIdx;
static std::vector<int> g_animPath;
static int g_animWinnerIdx = -1;
static double g_animDelay = 0.0;
static int g_blinkState = 0;     // remaining blink toggles
static int g_blinkIid = -1;

// Spin configuration
static int    CFG_SPIN_FAST_MS       = 18;
static int    CFG_SPIN_SLOW_MS       = 240;
static double CFG_SPIN_GROW          = 1.12;
static int    CFG_SPIN_ROUNDS        = 3;
static double CFG_SPIN_ROLLOUT_FACTOR = 1.5;
static int    CFG_BLINK_TIMES        = 3;
static int    CFG_BLINK_MS           = 180;

// Timer IDs
static constexpr UINT_PTR TIMER_ANIM  = 1001;
static constexpr UINT_PTR TIMER_BLINK = 1002;
static constexpr UINT_PTR TIMER_NEXT  = 1003;
static constexpr UINT_PTR TIMER_FINISH = 1004;

// Control IDs
static constexpr int IDC_LISTVIEW  = 2001;
static constexpr int IDC_BTN_LOAD  = 2002;
static constexpr int IDC_BTN_DRAW  = 2003;
static constexpr int IDC_BTN_CLEAR = 2004;
static constexpr int IDC_BTN_RELOAD = 2005;
static constexpr int IDC_EDIT_N    = 2006;

static constexpr int IDC_STATUSBAR = 2008;

// Random engine
static std::mt19937 g_rng(static_cast<unsigned>(std::time(nullptr)));

// ============================================================
//  Helpers: string conversions
// ============================================================
static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return {};
    int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    std::wstring w(n, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), &w[0], n);
    return w;
}

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string s(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &s[0], n, nullptr, nullptr);
    return s;
}

static std::wstring Trim(const std::wstring& s) {
    auto b = s.find_first_not_of(L" \t\r\n\xFEFF");  // also skip BOM
    if (b == std::wstring::npos) return {};
    auto e = s.find_last_not_of(L" \t\r\n");
    return s.substr(b, e - b + 1);
}

// ============================================================
//  CSV loading / saving
// ============================================================
static char DetectDelimiter(const std::string& firstLine) {
    int semicolons = 0, commas = 0;
    for (char c : firstLine) {
        if (c == ';') semicolons++;
        if (c == ',') commas++;
    }
    return (semicolons >= commas) ? ';' : ',';
}

static bool LoadCSV(const std::wstring& path) {
    std::ifstream ifs(WideToUtf8(path));
    if (!ifs.is_open()) {
        // Try wide path via _wfopen
        FILE* fp = _wfopen(path.c_str(), L"rb");
        if (!fp) return false;
        fseek(fp, 0, SEEK_END);
        long sz = ftell(fp);
        fseek(fp, 0, SEEK_SET);
        std::string buf(sz, '\0');
        fread(&buf[0], 1, sz, fp);
        fclose(fp);
        std::istringstream iss(buf);
        std::string line;
        // Detect delimiter from first line
        std::getline(iss, line);
        char delim = DetectDelimiter(line);
        // Skip header
        g_entries.clear();
        while (std::getline(iss, line)) {
            if (line.empty()) continue;
            std::istringstream ls(line);
            std::string nameField, counterField;
            std::getline(ls, nameField, delim);
            std::getline(ls, counterField, delim);
            std::wstring wname = Trim(Utf8ToWide(nameField));
            if (wname.empty() || wname == L"nan" || wname == L"None") continue;
            int cnt = 0;
            try { cnt = std::stoi(counterField); } catch (...) {}
            g_entries.push_back({wname, cnt});
        }
        return !g_entries.empty();
    }

    std::string line;
    std::getline(ifs, line);  // header
    char delim = DetectDelimiter(line);
    g_entries.clear();
    while (std::getline(ifs, line)) {
        if (line.empty()) continue;
        std::istringstream ls(line);
        std::string nameField, counterField;
        std::getline(ls, nameField, delim);
        std::getline(ls, counterField, delim);
        std::wstring wname = Trim(Utf8ToWide(nameField));
        if (wname.empty() || wname == L"nan" || wname == L"None") continue;
        int cnt = 0;
        try { cnt = std::stoi(counterField); } catch (...) {}
        g_entries.push_back({wname, cnt});
    }
    return !g_entries.empty();
}

static bool SaveCSV(const std::wstring& path) {
    // Write to temp first, then rename
    std::wstring tmpPath = path + L".tmp";
    FILE* fp = _wfopen(tmpPath.c_str(), L"wb");
    if (!fp) return false;

    // UTF-8 BOM for Excel compatibility
    fwrite("\xEF\xBB\xBF", 1, 3, fp);
    std::string header = "Name;Counter\r\n";
    fwrite(header.c_str(), 1, header.size(), fp);
    for (auto& e : g_entries) {
        std::string line = WideToUtf8(e.name) + ";" + std::to_string(e.counter) + "\r\n";
        fwrite(line.c_str(), 1, line.size(), fp);
    }
    fclose(fp);

    // Atomic-ish replace
    _wremove(path.c_str());
    if (_wrename(tmpPath.c_str(), path.c_str()) != 0) {
        // Fallback: copy tmp over
        CopyFileW(tmpPath.c_str(), path.c_str(), FALSE);
        _wremove(tmpPath.c_str());
    }
    return true;
}

// ============================================================
//  Fairness logic
// ============================================================
static std::vector<int> EligibleIndices() {
    if (g_entries.empty()) return {};
    int minC = g_entries[0].counter;
    for (auto& e : g_entries) minC = std::min(minC, e.counter);
    std::vector<int> result;
    for (int i = 0; i < (int)g_entries.size(); i++) {
        if (g_entries[i].counter == minC) result.push_back(i);
    }
    return result;
}

static void NormalizeCounters() {
    if (g_entries.empty()) return;
    int minC = g_entries[0].counter;
    for (auto& e : g_entries) minC = std::min(minC, e.counter);
    if (minC > 0) {
        for (auto& e : g_entries) e.counter -= minC;
    }
}

// ============================================================
//  ListView helpers
// ============================================================
static void SetStatus(const std::wstring& text) {
    SendMessageW(g_hStatus, SB_SETTEXTW, 0, (LPARAM)text.c_str());
}

static void PopulateListView() {
    ListView_DeleteAllItems(g_hList);
    for (int i = 0; i < (int)g_entries.size(); i++) {
        LVITEMW lvi = {};
        lvi.mask = LVIF_TEXT;
        lvi.iItem = i;
        lvi.iSubItem = 0;
        lvi.pszText = (LPWSTR)g_entries[i].name.c_str();
        ListView_InsertItem(g_hList, &lvi);

        std::wstring cnt = std::to_wstring(g_entries[i].counter);
        ListView_SetItemText(g_hList, i, 1, (LPWSTR)cnt.c_str());
    }
}

static void RefreshCounters() {
    for (int i = 0; i < (int)g_entries.size(); i++) {
        std::wstring cnt = std::to_wstring(g_entries[i].counter);
        ListView_SetItemText(g_hList, i, 1, (LPWSTR)cnt.c_str());
    }
}

static void HighlightRow(int idx, COLORREF bg) {
    // We use custom draw (NM_CUSTOMDRAW) for highlighting.
    // Store which row is highlighted and what color.
    // Trigger repaint.
    ListView_RedrawItems(g_hList, idx, idx);
    UpdateWindow(g_hList);
}

// Highlight state for custom draw
static int g_scanHighlightRow = -1;        // yellow scan highlight
static std::set<int> g_winnerRows;         // green winner rows

static void ClearScanHighlight() {
    if (g_scanHighlightRow >= 0) {
        int old = g_scanHighlightRow;
        g_scanHighlightRow = -1;
        ListView_RedrawItems(g_hList, old, old);
    }
}

static void SetScanHighlight(int idx) {
    ClearScanHighlight();
    g_scanHighlightRow = idx;
    ListView_RedrawItems(g_hList, idx, idx);
    ListView_EnsureVisible(g_hList, idx, FALSE);
    UpdateWindow(g_hList);
}

static void SetWinnerHighlight(int idx) {
    g_winnerRows.insert(idx);
    ListView_RedrawItems(g_hList, idx, idx);
    UpdateWindow(g_hList);
}

static void ClearWinnerHighlight(int idx) {
    g_winnerRows.erase(idx);
    ListView_RedrawItems(g_hList, idx, idx);
    UpdateWindow(g_hList);
}

static void ClearAllHighlights() {
    g_scanHighlightRow = -1;
    g_winnerRows.clear();
    InvalidateRect(g_hList, nullptr, TRUE);
    UpdateWindow(g_hList);
}

// ============================================================
//  Animation logic
// ============================================================
static void StopAllTimers() {
    KillTimer(g_hWnd, TIMER_ANIM);
    KillTimer(g_hWnd, TIMER_BLINK);
    KillTimer(g_hWnd, TIMER_NEXT);
    KillTimer(g_hWnd, TIMER_FINISH);
}

static void EndRoundEarly() {
    StopAllTimers();
    g_animRunning = false;
    g_roundSelectedIdx.clear();
    g_roundExcludedIdx.clear();
    EnableWindow(g_hBtnDraw, TRUE);
    SetStatus(L"Ziehung abgebrochen.");
}

static void FinishRound();
static void DrawNextOne();
static void ApplyWinnerAndContinue();
static void StartBlink();

static void OnAnimTimer() {
    if (g_animPath.empty()) {
        KillTimer(g_hWnd, TIMER_ANIM);
        // Animation done, show winner
        SetScanHighlight(g_animWinnerIdx);
        SetWinnerHighlight(g_animWinnerIdx);
        g_roundSelectedIdx.insert(g_animWinnerIdx);
        g_roundExcludedIdx.insert(g_animWinnerIdx);
        StartBlink();
        return;
    }
    int idx = g_animPath.front();
    g_animPath.erase(g_animPath.begin());
    SetScanHighlight(idx);

    g_animDelay = std::min(g_animDelay * CFG_SPIN_GROW, (double)CFG_SPIN_SLOW_MS);
    int ms = std::max(5, (int)g_animDelay);
    SetTimer(g_hWnd, TIMER_ANIM, ms, nullptr);
}

static void StartBlink() {
    g_blinkState = CFG_BLINK_TIMES * 2;
    g_blinkIid = g_animWinnerIdx;
    SetTimer(g_hWnd, TIMER_BLINK, CFG_BLINK_MS, nullptr);
}

static void OnBlinkTimer() {
    if (g_blinkState <= 0) {
        KillTimer(g_hWnd, TIMER_BLINK);
        SetWinnerHighlight(g_blinkIid);
        ClearScanHighlight();
        // Continue after short delay
        SetTimer(g_hWnd, TIMER_NEXT, 100, nullptr);
        return;
    }
    // Toggle winner highlight
    if (g_winnerRows.count(g_blinkIid)) {
        ClearWinnerHighlight(g_blinkIid);
        g_scanHighlightRow = -1;
        ListView_RedrawItems(g_hList, g_blinkIid, g_blinkIid);
    } else {
        SetWinnerHighlight(g_blinkIid);
    }
    g_blinkState--;
}

static void OnNextTimer() {
    KillTimer(g_hWnd, TIMER_NEXT);
    ApplyWinnerAndContinue();
}

static void ApplyWinnerAndContinue() {
    int idx = g_animWinnerIdx;
    g_entries[idx].counter++;
    NormalizeCounters();
    RefreshCounters();

    g_drawnCount++;
    std::wstring status = L"Gezogen: " + std::to_wstring(g_drawnCount) + L"/"
                        + std::to_wstring(g_toDrawTotal) + L" \u2013 Gewinner: "
                        + g_entries[idx].name;
    SetStatus(status);

    if (g_drawnCount < g_toDrawTotal) {
        SetTimer(g_hWnd, TIMER_FINISH, 300, nullptr);  // short pause then next
    } else {
        FinishRound();
    }
}

static void OnFinishTimer() {
    KillTimer(g_hWnd, TIMER_FINISH);
    DrawNextOne();
}

static void DrawNextOne() {
    if (g_entries.empty()) { EndRoundEarly(); return; }

    auto elig = EligibleIndices();
    if (elig.empty()) { NormalizeCounters(); elig = EligibleIndices(); }
    if (elig.empty()) { EndRoundEarly(); return; }

    // Exclude already-picked in this batch
    std::vector<int> filtered;
    for (int i : elig) {
        if (g_roundExcludedIdx.find(i) == g_roundExcludedIdx.end())
            filtered.push_back(i);
    }
    if (filtered.empty()) filtered = elig;  // all picked, allow repeats

    // Pick winner
    std::uniform_int_distribution<int> dist(0, (int)filtered.size() - 1);
    int winnerIdx = filtered[dist(g_rng)];
    g_animWinnerIdx = winnerIdx;

    // Build animation path
    g_animPath.clear();
    for (int r = 0; r < CFG_SPIN_ROUNDS; r++) {
        for (int i : filtered) g_animPath.push_back(i);
    }

    int rollout = std::max(1, (int)(filtered.size() * CFG_SPIN_ROLLOUT_FACTOR));
    std::uniform_int_distribution<int> startDist(0, (int)filtered.size() - 1);
    int cur = startDist(g_rng);
    int steps = 0;
    int maxSteps = rollout + (int)filtered.size() * 2;
    while (true) {
        int idx = filtered[cur];
        g_animPath.push_back(idx);
        steps++;
        if (steps >= rollout && idx == winnerIdx) break;
        if (steps >= maxSteps) { g_animPath.push_back(winnerIdx); break; }
        cur = (cur + 1) % (int)filtered.size();
    }

    g_animDelay = (double)CFG_SPIN_FAST_MS;
    SetTimer(g_hWnd, TIMER_ANIM, CFG_SPIN_FAST_MS, nullptr);
}

static void FinishRound() {
    // Save CSV
    if (!g_filePath.empty()) {
        if (!SaveCSV(g_filePath)) {
            MessageBoxW(g_hWnd, L"Fehler beim Speichern der Datei.", L"Speicherfehler", MB_OK | MB_ICONERROR);
        }
    }

    // Build summary
    std::wstring summary = L"Runde beendet \u2013 Gewinner: ";
    bool first = true;
    for (int i : g_roundSelectedIdx) {
        if (!first) summary += L", ";
        summary += g_entries[i].name;
        first = false;
    }
    SetStatus(summary);

    g_animRunning = false;
    EnableWindow(g_hBtnDraw, TRUE);
    g_roundSelectedIdx.clear();
    g_roundExcludedIdx.clear();
}

static void OnDrawClicked() {
    if (g_entries.empty() || g_animRunning) return;

    wchar_t buf[16] = {};
    GetWindowTextW(g_hEditN, buf, 15);
    int n = _wtoi(buf);
    if (n < 1) {
        MessageBoxW(g_hWnd, L"Bitte eine gültige Zahl eingeben.", L"Eingabe", MB_OK | MB_ICONWARNING);
        return;
    }
    n = std::min(n, (int)g_entries.size());

    g_toDrawTotal = n;
    g_drawnCount = 0;
    g_roundSelectedIdx.clear();
    g_roundExcludedIdx.clear();
    g_animRunning = true;
    EnableWindow(g_hBtnDraw, FALSE);

    SetStatus((L"Ziehe " + std::to_wstring(n) + L" Person(en) \u2026").c_str());
    DrawNextOne();
}

// ============================================================
//  File loading
// ============================================================
static void DoLoadFile(const std::wstring& path) {
    if (!LoadCSV(path)) {
        MessageBoxW(g_hWnd, L"Keine gültigen Namen in der Datei gefunden.\n\n"
                    L"Erwartetes Format: CSV (Semikolon oder Komma getrennt)\n"
                    L"Spalte 1: Name, Spalte 2: Counter (optional)",
                    L"Fehler beim Laden", MB_OK | MB_ICONERROR);
        return;
    }
    g_filePath = path;
    PopulateListView();
    EnableWindow(g_hBtnDraw, TRUE);
    EnableWindow(g_hBtnClear, TRUE);
    EnableWindow(g_hBtnReload, TRUE);

    // Extract filename for display
    std::wstring fname = path;
    auto pos = fname.find_last_of(L"\\/");
    if (pos != std::wstring::npos) fname = fname.substr(pos + 1);

    SetStatus((L"Geladen: " + std::to_wstring(g_entries.size()) + L" Einträge aus " + fname).c_str());
}

static void OnLoadExcel() {
    if (g_animRunning) {
        MessageBoxW(g_hWnd, L"Bitte warten, bis die aktuelle Ziehung beendet ist.",
                     L"Bitte warten", MB_OK | MB_ICONWARNING);
        return;
    }

    wchar_t szFile[MAX_PATH] = {};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hWnd;
    ofn.lpstrFilter = L"CSV-Dateien (*.csv)\0*.csv\0Alle Dateien (*.*)\0*.*\0";
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrTitle = L"CSV mit Namensliste auswählen";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

    if (GetOpenFileNameW(&ofn)) {
        DoLoadFile(szFile);
    }
}

static void OnReloadExcel() {
    if (g_animRunning) {
        MessageBoxW(g_hWnd, L"Bitte warten, bis die aktuelle Ziehung beendet ist.",
                     L"Bitte warten", MB_OK | MB_ICONWARNING);
        return;
    }
    if (g_filePath.empty()) return;
    DoLoadFile(g_filePath);
}

// ============================================================
//  Settings dialog (manual modal window approach)
// ============================================================
static void ShowConfigDialog() {
    // We'll create a simple window manually instead of a dialog template
    // for simplicity and to avoid resource file dependency.
    // Using a modal dialog with CreateDialogIndirect.

    // Actually, let's use a simple approach: create a popup window with controls.
    HWND hDlg = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"#32770",  // dialog class
        L"Spin-Parameter",
        WS_VISIBLE | WS_SYSMENU | WS_CAPTION | DS_MODALFRAME | WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT, 380, 340,
        g_hWnd, nullptr, g_hInst, nullptr
    );
    if (!hDlg) return;

    // We'll use a simpler approach: just a message box with info,
    // or build a proper child-window dialog.
    // For robustness, let's build it with CreateWindow children.

    EnableWindow(g_hWnd, FALSE);  // modal

    struct Field { const wchar_t* label; int editId; const wchar_t* value; };
    wchar_t buf[7][32];
    swprintf(buf[0], 32, L"%d", CFG_SPIN_FAST_MS);
    swprintf(buf[1], 32, L"%d", CFG_SPIN_SLOW_MS);
    swprintf(buf[2], 32, L"%.2f", CFG_SPIN_GROW);
    swprintf(buf[3], 32, L"%d", CFG_SPIN_ROUNDS);
    swprintf(buf[4], 32, L"%.2f", CFG_SPIN_ROLLOUT_FACTOR);
    swprintf(buf[5], 32, L"%d", CFG_BLINK_TIMES);
    swprintf(buf[6], 32, L"%d", CFG_BLINK_MS);

    Field fields[] = {
        {L"Startgeschwindigkeit (ms):", 3001, buf[0]},
        {L"Endgeschwindigkeit (ms):",   3002, buf[1]},
        {L"Verzögerungsfaktor (>1.0):", 3003, buf[2]},
        {L"Spin-Runden:",               3004, buf[3]},
        {L"Ausroll-Faktor:",            3005, buf[4]},
        {L"Blinkanzahl:",               3006, buf[5]},
        {L"Blinktempo (ms):",           3007, buf[6]},
    };

    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    int y = 12;
    for (auto& f : fields) {
        HWND hLbl = CreateWindowW(L"STATIC", f.label, WS_CHILD | WS_VISIBLE,
                                   12, y + 2, 210, 20, hDlg, nullptr, g_hInst, nullptr);
        HWND hEdit = CreateWindowW(L"EDIT", f.value,
                                    WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
                                    230, y, 80, 22, hDlg, (HMENU)(INT_PTR)f.editId, g_hInst, nullptr);
        SendMessageW(hLbl, WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessageW(hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
        y += 30;
    }

    // Buttons
    HWND hOK = CreateWindowW(L"BUTTON", L"Übernehmen",
                              WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                              140, y + 10, 100, 28, hDlg, (HMENU)IDOK, g_hInst, nullptr);
    HWND hCancel = CreateWindowW(L"BUTTON", L"Abbrechen",
                                  WS_CHILD | WS_VISIBLE,
                                  250, y + 10, 90, 28, hDlg, (HMENU)IDCANCEL, g_hInst, nullptr);
    SendMessageW(hOK, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessageW(hCancel, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Center on parent
    RECT rc, rp;
    GetWindowRect(hDlg, &rc);
    GetWindowRect(g_hWnd, &rp);
    int cx = (rp.left + rp.right) / 2 - (rc.right - rc.left) / 2;
    int cy = (rp.top + rp.bottom) / 2 - (rc.bottom - rc.top) / 2;
    SetWindowPos(hDlg, HWND_TOP, cx, cy, 0, 0, SWP_NOSIZE);

    // Message loop for this modal dialog
    MSG msg;
    bool running = true;
    while (running && GetMessageW(&msg, nullptr, 0, 0)) {
        if (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd)) {
            if (msg.message == WM_COMMAND) {
                int id = LOWORD(msg.wParam);
                if (id == IDOK) {
                    // Read values
                    auto getInt = [&](int eid, int minV) -> int {
                        wchar_t b[32]; GetDlgItemTextW(hDlg, eid, b, 32);
                        return std::max(minV, _wtoi(b));
                    };
                    auto getDbl = [&](int eid, double minV) -> double {
                        wchar_t b[32]; GetDlgItemTextW(hDlg, eid, b, 32);
                        return std::max(minV, _wtof(b));
                    };
                    CFG_SPIN_FAST_MS = getInt(3001, 5);
                    CFG_SPIN_SLOW_MS = std::max(CFG_SPIN_FAST_MS, getInt(3002, 5));
                    CFG_SPIN_GROW = getDbl(3003, 1.01);
                    CFG_SPIN_ROUNDS = getInt(3004, 0);
                    CFG_SPIN_ROLLOUT_FACTOR = getDbl(3005, 0.0);
                    CFG_BLINK_TIMES = getInt(3006, 0);
                    CFG_BLINK_MS = getInt(3007, 20);
                    running = false;
                } else if (id == IDCANCEL) {
                    running = false;
                }
                continue;
            }
            if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE) {
                running = false;
                continue;
            }
        }
        if (msg.message == WM_CLOSE && msg.hwnd == hDlg) {
            running = false;
            continue;
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    EnableWindow(g_hWnd, TRUE);
    DestroyWindow(hDlg);
    SetForegroundWindow(g_hWnd);
}

// ============================================================
//  Main window
// ============================================================
static void CreateMainControls(HWND hWnd) {
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    // Top bar
    g_hBtnLoad = CreateWindowW(L"BUTTON", L"CSV laden \u2026",
        WS_CHILD | WS_VISIBLE, 8, 8, 110, 28, hWnd, (HMENU)IDC_BTN_LOAD, g_hInst, nullptr);

    // ListView
    g_hList = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        WC_LISTVIEWW, L"",
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_NOSORTHEADER | LVS_SHOWSELALWAYS,
        8, 44, 700, 440, hWnd, (HMENU)IDC_LISTVIEW, g_hInst, nullptr);
    ListView_SetExtendedListViewStyle(g_hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

    // Columns
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
    col.pszText = (LPWSTR)L"Name";
    col.cx = 530;
    col.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(g_hList, 0, &col);

    col.pszText = (LPWSTR)L"Gezogen";
    col.cx = 100;
    col.fmt = LVCFMT_CENTER;
    ListView_InsertColumn(g_hList, 1, &col);

    // Control bar
    int cy = 492;
    CreateWindowW(L"STATIC", L"Anzahl ziehen:", WS_CHILD | WS_VISIBLE,
                   8, cy + 4, 100, 20, hWnd, nullptr, g_hInst, nullptr);

    g_hEditN = CreateWindowW(L"EDIT", L"1",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_AUTOHSCROLL,
        112, cy, 60, 24, hWnd, (HMENU)IDC_EDIT_N, g_hInst, nullptr);

    g_hBtnDraw = CreateWindowW(L"BUTTON", L"Ziehung starten",
        WS_CHILD | WS_VISIBLE | WS_DISABLED, 180, cy, 130, 28, hWnd, (HMENU)IDC_BTN_DRAW, g_hInst, nullptr);

    g_hBtnClear = CreateWindowW(L"BUTTON", L"Markierungen zurücksetzen",
        WS_CHILD | WS_VISIBLE | WS_DISABLED, 320, cy, 190, 28, hWnd, (HMENU)IDC_BTN_CLEAR, g_hInst, nullptr);

    g_hBtnReload = CreateWindowW(L"BUTTON", L"\u27F3 Neu laden",
        WS_CHILD | WS_VISIBLE | WS_DISABLED, 520, cy, 100, 28, hWnd, (HMENU)IDC_BTN_RELOAD, g_hInst, nullptr);

    // Status bar
    g_hStatus = CreateWindowW(STATUSCLASSNAMEW, L"Bereit.",
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
        0, 0, 0, 0, hWnd, (HMENU)IDC_STATUSBAR, g_hInst, nullptr);

    // Set font on all controls
    HWND controls[] = {g_hBtnLoad, g_hBtnDraw, g_hBtnClear, g_hBtnReload, g_hEditN, g_hList};
    for (HWND h : controls) {
        SendMessageW(h, WM_SETFONT, (WPARAM)hFont, TRUE);
    }
    // Also set font on static labels
    EnumChildWindows(hWnd, [](HWND h, LPARAM lParam) -> BOOL {
        SendMessageW(h, WM_SETFONT, (WPARAM)lParam, TRUE);
        return TRUE;
    }, (LPARAM)hFont);
}

static void OnResize(HWND hWnd) {
    RECT rc;
    GetClientRect(hWnd, &rc);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;

    // Status bar auto-resizes
    SendMessageW(g_hStatus, WM_SIZE, 0, 0);
    RECT sbar;
    GetWindowRect(g_hStatus, &sbar);
    int statusH = sbar.bottom - sbar.top;

    int controlY = h - statusH - 40;
    int listH = controlY - 52;

    MoveWindow(g_hList, 8, 44, w - 16, std::max(100, listH), TRUE);

    // Adjust name column width
    int nameColW = w - 16 - 110 - 20;  // subtract counter col + scrollbar
    ListView_SetColumnWidth(g_hList, 0, std::max(100, nameColW));

    // Control bar
    MoveWindow(g_hBtnDraw, 180, controlY, 130, 28, TRUE);
    MoveWindow(g_hBtnClear, 320, controlY, 190, 28, TRUE);
    MoveWindow(g_hBtnReload, 520, controlY, std::min(120, w - 530), 28, TRUE);
    MoveWindow(g_hEditN, 112, controlY, 50, 24, TRUE);

    // Reposition static label
    HWND hLabel = FindWindowExW(hWnd, nullptr, L"Static", L"Anzahl ziehen:");
    if (hLabel) MoveWindow(hLabel, 8, controlY + 4, 100, 20, TRUE);

    InvalidateRect(hWnd, nullptr, TRUE);
}

// Build menu bar
static void BuildMenu(HWND hWnd) {
    HMENU hMenu = CreateMenu();
    HMENU hSettings = CreatePopupMenu();
    AppendMenuW(hSettings, MF_STRING, 4001, L"Spin-Parameter \u2026");
    AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hSettings, L"Einstellungen");

    HMENU hHelp = CreatePopupMenu();
    AppendMenuW(hHelp, MF_STRING, 4002, L"Über \u2026");
    AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hHelp, L"Hilfe");

    SetMenu(hWnd, hMenu);
}

// ============================================================
//  Window Procedure
// ============================================================
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        CreateMainControls(hWnd);
        BuildMenu(hWnd);
        return 0;

    case WM_SIZE:
        OnResize(hWnd);
        return 0;

    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id == IDC_BTN_LOAD) OnLoadExcel();
        else if (id == IDC_BTN_DRAW) OnDrawClicked();
        else if (id == IDC_BTN_CLEAR) { ClearAllHighlights(); SetStatus(L"Markierungen zurückgesetzt."); }
        else if (id == IDC_BTN_RELOAD) OnReloadExcel();
        else if (id == 4001) ShowConfigDialog();
        else if (id == 4002) {
            MessageBoxW(hWnd,
                L"Glücksrad – faire Zufallsauswahl (C++)\n\n"
                L"Spin-Dynamik, Blinken, CSV-Speicherung.\n"
                L"Nativ kompiliert für schnellen Start.",
                L"Über", MB_OK | MB_ICONINFORMATION);
        }
        return 0;
    }

    case WM_TIMER:
        if (wParam == TIMER_ANIM) OnAnimTimer();
        else if (wParam == TIMER_BLINK) OnBlinkTimer();
        else if (wParam == TIMER_NEXT) OnNextTimer();
        else if (wParam == TIMER_FINISH) OnFinishTimer();
        return 0;

    case WM_NOTIFY: {
        NMHDR* nmh = (NMHDR*)lParam;
        if (nmh->idFrom == IDC_LISTVIEW && nmh->code == NM_CUSTOMDRAW) {
            NMLVCUSTOMDRAW* lvcd = (NMLVCUSTOMDRAW*)lParam;
            switch (lvcd->nmcd.dwDrawStage) {
            case CDDS_PREPAINT:
                return CDRF_NOTIFYITEMDRAW;
            case CDDS_ITEMPREPAINT: {
                int row = (int)lvcd->nmcd.dwItemSpec;
                if (row == g_scanHighlightRow) {
                    lvcd->clrTextBk = RGB(255, 224, 130);  // yellow
                    lvcd->clrText = RGB(0, 0, 0);
                } else if (g_winnerRows.count(row)) {
                    lvcd->clrTextBk = RGB(200, 230, 201);  // green
                    lvcd->clrText = RGB(0, 0, 0);
                } else {
                    lvcd->clrTextBk = RGB(255, 255, 255);
                    lvcd->clrText = RGB(0, 0, 0);
                }
                return CDRF_NEWFONT;
            }
            }
        }
        break;
    }

    case WM_CLOSE:
        StopAllTimers();
        DestroyWindow(hWnd);
        return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ============================================================
//  WinMain
// ============================================================
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    g_hInst = hInstance;

    // Init common controls (for ListView, StatusBar, UpDown)
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES };
    InitCommonControlsEx(&icc);

    // Register window class
    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"GluecksradClass";
    wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
    RegisterClassExW(&wc);

    // Create main window
    g_hWnd = CreateWindowExW(
        0, L"GluecksradClass",
        L"Glücksrad \u2013 Faire Zufallsauswahl",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 760, 600,
        nullptr, nullptr, hInstance, nullptr
    );

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    // Message loop
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return (int)msg.wParam;
}

// Console subsystem fallback (for testing with g++ without -mwindows)
#ifndef _MSC_VER
int wmain(int, wchar_t**) {
    return wWinMain(GetModuleHandleW(nullptr), nullptr, nullptr, SW_SHOW);
}
int main() {
    return wWinMain(GetModuleHandleW(nullptr), nullptr, nullptr, SW_SHOW);
}
#endif
