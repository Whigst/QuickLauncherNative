#include "LauncherApp.h"

#include <commctrl.h>
#include <shlobj.h>

#include <fstream>
#include <memory>
#include <sstream>
#include <unordered_set>
#include <utility>

namespace quicklauncher
{
namespace
{
std::string Utf8FromWide(const std::wstring& value)
{
    if (value.empty())
    {
        return {};
    }

    const int size = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    std::string result(size, '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), result.data(), size, nullptr, nullptr);
    return result;
}

std::wstring WideFromUtf8(const std::string& value)
{
    if (value.empty())
    {
        return {};
    }

    const int size = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
    std::wstring result(size, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), result.data(), size);
    return result;
}

void WriteUtf8Line(std::ofstream& stream, const std::wstring& line)
{
    const auto utf8 = Utf8FromWide(line);
    stream.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
    stream.put('\n');
}

std::vector<std::wstring> GetFixedDriveRoots()
{
    std::vector<std::wstring> roots;
    const DWORD mask = GetLogicalDrives();

    for (wchar_t drive = L'A'; drive <= L'Z'; ++drive)
    {
        if ((mask & (1u << (drive - L'A'))) == 0)
        {
            continue;
        }

        std::wstring root;
        root += drive;
        root += L":\\";

        if (GetDriveTypeW(root.c_str()) == DRIVE_FIXED)
        {
            roots.push_back(root);
        }
    }

    return roots;
}

std::wstring TrimCopy(const std::wstring& value)
{
    const auto begin = value.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos)
    {
        return L"";
    }

    const auto end = value.find_last_not_of(L" \t\r\n");
    return value.substr(begin, end - begin + 1);
}

std::wstring NormalizeDirectoryPath(std::wstring path)
{
    std::replace(path.begin(), path.end(), L'/', L'\\');
    path = ToLowerCopy(TrimCopy(path));

    while (path.size() > 3 && !path.empty() && (path.back() == L'\\' || path.back() == L'/'))
    {
        path.pop_back();
    }

    return path;
}

std::wstring StripWrappingQuotes(std::wstring value)
{
    value = TrimCopy(value);
    if (value.size() >= 2 && value.front() == L'"' && value.back() == L'"')
    {
        return value.substr(1, value.size() - 2);
    }

    return value;
}

bool IsSameOrChildPath(const std::wstring& path, const std::wstring& basePath)
{
    if (basePath.empty() || path.size() < basePath.size())
    {
        return false;
    }

    if (path.compare(0, basePath.size(), basePath) != 0)
    {
        return false;
    }

    if (path.size() == basePath.size())
    {
        return true;
    }

    if (!basePath.empty() && basePath.back() == L'\\')
    {
        return true;
    }

    return path[basePath.size()] == L'\\';
}

std::wstring BuildSourceTag(const std::wstring& fullPath, bool isDirectory)
{
    const auto container = isDirectory ? GetDirectoryPart(fullPath) : GetDirectoryPart(fullPath);
    auto source = GetLeafName(container);

    if (source.empty())
    {
        source = container;
    }

    if (source.empty())
    {
        source = fullPath;
    }

    return source;
}

std::wstring BuildDisplayLabel(const LaunchEntry& entry)
{
    std::wstring label = entry.name;
    if (!entry.isDirectory)
    {
        label += entry.extension;
    }
    else
    {
        label += L" [Folder]";
    }

    if (!entry.sourceTag.empty())
    {
        label += L"  @";
        label += entry.sourceTag;
    }

    return label;
}

void TrimBucket(QueryBucket& bucket)
{
    if (bucket.launchCounts.size() <= 10)
    {
        return;
    }

    std::vector<std::pair<std::wstring, int>> pairs(bucket.launchCounts.begin(), bucket.launchCounts.end());
    std::sort(pairs.begin(), pairs.end(), [](const auto& left, const auto& right)
    {
        if (left.second != right.second)
        {
            return left.second > right.second;
        }

        return left.first < right.first;
    });

    bucket.launchCounts.clear();
    for (size_t index = 0; index < std::min<size_t>(10, pairs.size()); ++index)
    {
        bucket.launchCounts[pairs[index].first] = pairs[index].second;
    }
}

class ScopedCriticalSection
{
public:
    explicit ScopedCriticalSection(CRITICAL_SECTION& criticalSection)
        : criticalSection_(criticalSection)
    {
        EnterCriticalSection(&criticalSection_);
    }

    ~ScopedCriticalSection()
    {
        LeaveCriticalSection(&criticalSection_);
    }

private:
    CRITICAL_SECTION& criticalSection_;
};

struct IndexThreadArgs
{
    LauncherApp* app = nullptr;
    bool forceRebuild = false;
};

constexpr DWORD kWatchNotifyFilter =
    FILE_NOTIFY_CHANGE_FILE_NAME
    | FILE_NOTIFY_CHANGE_DIR_NAME
    | FILE_NOTIFY_CHANGE_LAST_WRITE;

void AppendUniqueRoot(std::vector<std::wstring>& roots, std::wstring candidate)
{
    candidate = NormalizeDirectoryPath(std::move(candidate));
    if (candidate.empty())
    {
        return;
    }

    for (auto it = roots.begin(); it != roots.end();)
    {
        if (IsSameOrChildPath(candidate, *it))
        {
            return;
        }

        if (IsSameOrChildPath(*it, candidate))
        {
            it = roots.erase(it);
            continue;
        }

        ++it;
    }

    roots.push_back(std::move(candidate));
}

struct HotkeyParseResult
{
    UINT modifiers = MOD_ALT;
    UINT virtualKey = VK_SPACE;
    std::wstring label = L"Alt+Space";
};

std::vector<std::wstring> SplitHotkeyTokens(const std::wstring& value)
{
    std::vector<std::wstring> tokens;
    std::wstring current;

    for (const auto ch : value)
    {
        if (ch == L'+')
        {
            tokens.push_back(TrimCopy(current));
            current.clear();
            continue;
        }

        current.push_back(ch);
    }

    tokens.push_back(TrimCopy(current));
    return tokens;
}

UINT ParseHotkeyKeyToken(const std::wstring& token)
{
    if (token.size() == 1)
    {
        const wchar_t ch = token[0];
        if (ch >= L'a' && ch <= L'z')
        {
            return static_cast<UINT>(ch - (L'a' - L'A'));
        }

        if ((ch >= L'A' && ch <= L'Z') || (ch >= L'0' && ch <= L'9'))
        {
            return static_cast<UINT>(ch);
        }
    }

    if (token.size() >= 2 && token[0] == L'f')
    {
        const int functionIndex = _wtoi(token.substr(1).c_str());
        if (functionIndex >= 1 && functionIndex <= 24)
        {
            return static_cast<UINT>(VK_F1 + functionIndex - 1);
        }
    }

    if (token == L"space" || token == L"spacebar")
    {
        return VK_SPACE;
    }

    if (token == L"tab")
    {
        return VK_TAB;
    }

    if (token == L"enter" || token == L"return")
    {
        return VK_RETURN;
    }

    if (token == L"esc" || token == L"escape")
    {
        return VK_ESCAPE;
    }

    if (token == L"backspace" || token == L"bs")
    {
        return VK_BACK;
    }

    if (token == L"delete" || token == L"del")
    {
        return VK_DELETE;
    }

    if (token == L"insert" || token == L"ins")
    {
        return VK_INSERT;
    }

    if (token == L"home")
    {
        return VK_HOME;
    }

    if (token == L"end")
    {
        return VK_END;
    }

    if (token == L"pageup" || token == L"pgup")
    {
        return VK_PRIOR;
    }

    if (token == L"pagedown" || token == L"pgdn")
    {
        return VK_NEXT;
    }

    if (token == L"up")
    {
        return VK_UP;
    }

    if (token == L"down")
    {
        return VK_DOWN;
    }

    if (token == L"left")
    {
        return VK_LEFT;
    }

    if (token == L"right")
    {
        return VK_RIGHT;
    }

    return 0;
}

std::wstring HotkeyKeyLabel(UINT virtualKey)
{
    if ((virtualKey >= L'A' && virtualKey <= L'Z') || (virtualKey >= L'0' && virtualKey <= L'9'))
    {
        return std::wstring(1, static_cast<wchar_t>(virtualKey));
    }

    if (virtualKey >= VK_F1 && virtualKey <= VK_F24)
    {
        return L"F" + std::to_wstring(virtualKey - VK_F1 + 1);
    }

    switch (virtualKey)
    {
    case VK_SPACE:
        return L"Space";
    case VK_TAB:
        return L"Tab";
    case VK_RETURN:
        return L"Enter";
    case VK_ESCAPE:
        return L"Esc";
    case VK_BACK:
        return L"Backspace";
    case VK_DELETE:
        return L"Delete";
    case VK_INSERT:
        return L"Insert";
    case VK_HOME:
        return L"Home";
    case VK_END:
        return L"End";
    case VK_PRIOR:
        return L"PageUp";
    case VK_NEXT:
        return L"PageDown";
    case VK_UP:
        return L"Up";
    case VK_DOWN:
        return L"Down";
    case VK_LEFT:
        return L"Left";
    case VK_RIGHT:
        return L"Right";
    default:
        return L"Key";
    }
}

std::wstring BuildHotkeyLabel(UINT modifiers, UINT virtualKey)
{
    std::wstring label;

    auto appendPart = [&label](const std::wstring& value)
    {
        if (!label.empty())
        {
            label += L"+";
        }

        label += value;
    };

    if ((modifiers & MOD_CONTROL) != 0)
    {
        appendPart(L"Ctrl");
    }

    if ((modifiers & MOD_ALT) != 0)
    {
        appendPart(L"Alt");
    }

    if ((modifiers & MOD_SHIFT) != 0)
    {
        appendPart(L"Shift");
    }

    if ((modifiers & MOD_WIN) != 0)
    {
        appendPart(L"Win");
    }

    appendPart(HotkeyKeyLabel(virtualKey));
    return label;
}

bool TryParseHotkey(const std::wstring& value, HotkeyParseResult& result)
{
    UINT modifiers = 0;
    UINT virtualKey = 0;

    for (const auto& tokenRaw : SplitHotkeyTokens(value))
    {
        const auto token = ToLowerCopy(tokenRaw);
        if (token.empty())
        {
            return false;
        }

        if (token == L"ctrl" || token == L"control")
        {
            modifiers |= MOD_CONTROL;
            continue;
        }

        if (token == L"alt")
        {
            modifiers |= MOD_ALT;
            continue;
        }

        if (token == L"shift")
        {
            modifiers |= MOD_SHIFT;
            continue;
        }

        if (token == L"win" || token == L"windows" || token == L"cmd")
        {
            modifiers |= MOD_WIN;
            continue;
        }

        if (virtualKey != 0)
        {
            return false;
        }

        virtualKey = ParseHotkeyKeyToken(token);
        if (virtualKey == 0)
        {
            return false;
        }
    }

    if (modifiers == 0 || virtualKey == 0)
    {
        return false;
    }

    result.modifiers = modifiers;
    result.virtualKey = virtualKey;
    result.label = BuildHotkeyLabel(modifiers, virtualKey);
    return true;
}

bool IsPreferredLaunchSurface(const LaunchEntry& entry)
{
    return !entry.isDirectory
        && !entry.isUrl
        && (entry.extension == L".lnk"
            || entry.extension == L".appref-ms"
            || entry.sourceTag == L"Start Menu"
            || entry.sourceTag == L"Desktop"
            || entry.sourceTag == L"Pinned");
}

bool IsScriptLikeExtension(const std::wstring& extension)
{
    return extension == L".ps1"
        || extension == L".vbs"
        || extension == L".vbe"
        || extension == L".js"
        || extension == L".jse"
        || extension == L".wsf";
}

bool IsPrimaryLaunchExtension(const std::wstring& extension)
{
    return extension == L".exe"
        || extension == L".lnk"
        || extension == L".appref-ms"
        || extension == L".bat"
        || extension == L".cmd"
        || extension == L".com"
        || extension == L".msc"
        || extension == L".msi";
}

int GetEntryResultRank(const LaunchEntry& entry)
{
    if (IsPreferredLaunchSurface(entry))
    {
        return 0;
    }

    if (!entry.isDirectory && !entry.isUrl && IsPrimaryLaunchExtension(entry.extension))
    {
        return 1;
    }

    if (!entry.isDirectory && !entry.isUrl)
    {
        return 2;
    }

    if (entry.isDirectory)
    {
        return 3;
    }

    return 4;
}

int GetEntryCategoryBoost(const LaunchEntry& entry, size_t queryLength)
{
    if (entry.isUrl)
    {
        return -240;
    }

    if (entry.isDirectory)
    {
        return queryLength <= 3 ? -420 : -260;
    }

    if (IsPreferredLaunchSurface(entry))
    {
        return queryLength <= 3 ? 1800 : 1100;
    }

    if (IsPrimaryLaunchExtension(entry.extension))
    {
        return queryLength <= 3 ? 900 : 520;
    }

    if (IsScriptLikeExtension(entry.extension))
    {
        return queryLength <= 3 ? -320 : -140;
    }

    return 120;
}

void AppendPrefixKeys(std::unordered_set<std::wstring>& keys, const std::wstring& token)
{
    if (token.empty())
    {
        return;
    }

    keys.insert(token.substr(0, 1));
    if (token.size() >= 2)
    {
        keys.insert(token.substr(0, 2));
    }
}

void AppendIndexedCandidates(
    std::vector<size_t>& candidates,
    std::unordered_set<size_t>& seen,
    const std::unordered_map<std::wstring, std::vector<size_t>>& prefixIndex,
    const std::wstring& key)
{
    if (key.empty())
    {
        return;
    }

    const auto bucket = prefixIndex.find(key);
    if (bucket == prefixIndex.end())
    {
        return;
    }

    candidates.reserve(candidates.size() + bucket->second.size());
    for (const auto entryIndex : bucket->second)
    {
        if (seen.insert(entryIndex).second)
        {
            candidates.push_back(entryIndex);
        }
    }
}
}  // namespace

LauncherApp::~LauncherApp()
{
    shuttingDown_ = true;
    StopSearchThread();
    StopWatchThread();

    Shell_NotifyIconW(NIM_DELETE, &trayIconData_);
    UnregisterLauncherHotkey();

    if (indexThread_ != nullptr)
    {
        CloseHandle(indexThread_);
        indexThread_ = nullptr;
    }

    if (titleFont_ != nullptr)
    {
        DeleteObject(titleFont_);
    }

    if (normalFont_ != nullptr)
    {
        DeleteObject(normalFont_);
    }

    if (smallFont_ != nullptr)
    {
        DeleteObject(smallFont_);
    }

    DeleteCriticalSection(&dataLock_);
}

bool LauncherApp::Initialize(HINSTANCE instance)
{
    instance_ = instance;
    InitializeCriticalSection(&dataLock_);

    INITCOMMONCONTROLSEX commonControls{};
    commonControls.dwSize = sizeof(commonControls);
    commonControls.dwICC = ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&commonControls);

    dataDirectory_ = GetLocalAppDataPath() + L"\\QuickLauncherNative";
    cachePath_ = dataDirectory_ + L"\\launcher-index-v3.tsv";
    usagePath_ = dataDirectory_ + L"\\launch-usage-v1.tsv";
    settingsPath_ = dataDirectory_ + L"\\settings.ini";
    CreateDirectoryW(dataDirectory_.c_str(), nullptr);

    wchar_t userProfileBuffer[MAX_PATH * 4] = {};
    if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_PROFILE, nullptr, SHGFP_TYPE_CURRENT, userProfileBuffer)))
    {
        userProfilePath_ = NormalizeDirectoryPath(userProfileBuffer);
    }
    InitializePreferredRoots();

    WNDCLASSEXW windowClass{};
    windowClass.cbSize = sizeof(windowClass);
    windowClass.lpfnWndProc = WindowProc;
    windowClass.style = CS_DROPSHADOW;
    windowClass.hInstance = instance_;
    windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    windowClass.lpszClassName = L"QuickLauncherNativeWindow";

    if (RegisterClassExW(&windowClass) == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
    {
        return false;
    }

    mainWindow_ = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
        windowClass.lpszClassName,
        L"QuickLauncherNative",
        WS_POPUP | WS_CLIPCHILDREN,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        880,
        382,
        nullptr,
        nullptr,
        instance_,
        this);

    if (mainWindow_ == nullptr)
    {
        return false;
    }

    autostartEnabled_ = IsAutostartEnabled();
    EnsureSettingsFileExists();
    LoadSettings();
    LoadUsage();
    LoadCache();
    CreateTrayIcon();
    RegisterLauncherHotkey();
    StartWatchThread();
    StartSearchThread();
    UpdateResults();
    UpdateFooter();

    if (entries_ == nullptr || entries_->empty())
    {
        StartIndexThread(true);
    }
    else if (CurrentUtcTicks() - lastIndexedUtcTicks_ > 6LL * kTicksPerHour)
    {
        StartIndexThread(false);
    }

    return true;
}

int LauncherApp::Run()
{
    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }

    return static_cast<int>(message.wParam);
}

LRESULT CALLBACK LauncherApp::WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
    LauncherApp* app = nullptr;

    if (message == WM_NCCREATE)
    {
        auto* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
        app = static_cast<LauncherApp*>(createStruct->lpCreateParams);
        app->mainWindow_ = window;
        SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(app));
    }
    else
    {
        app = reinterpret_cast<LauncherApp*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    }

    if (app == nullptr)
    {
        return DefWindowProcW(window, message, wParam, lParam);
    }

    return app->HandleMessage(window, message, wParam, lParam);
}

LRESULT CALLBACK LauncherApp::EditSubclassProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData)
{
    auto* app = reinterpret_cast<LauncherApp*>(refData);
    return app->HandleEditMessage(window, message, wParam, lParam);
}

LRESULT CALLBACK LauncherApp::ListSubclassProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData)
{
    auto* app = reinterpret_cast<LauncherApp*>(refData);
    return app->HandleListMessage(window, message, wParam, lParam);
}

DWORD WINAPI LauncherApp::IndexThreadProc(LPVOID param)
{
    std::unique_ptr<IndexThreadArgs> args(static_cast<IndexThreadArgs*>(param));
    args->app->IndexWorker(args->forceRebuild);
    return 0;
}

DWORD WINAPI LauncherApp::WatchThreadProc(LPVOID param)
{
    auto* app = static_cast<LauncherApp*>(param);
    app->WatchLoop();
    return 0;
}

DWORD WINAPI LauncherApp::SearchThreadProc(LPVOID param)
{
    auto* app = static_cast<LauncherApp*>(param);
    app->SearchLoop();
    return 0;
}

LRESULT LauncherApp::HandleMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        CreateFonts();
        CreateControls();
        LayoutControls();
        return 0;

    case WM_SIZE:
        LayoutControls();
        InvalidateRect(mainWindow_, nullptr, TRUE);
        return 0;

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
    {
        PAINTSTRUCT paint{};
        const HDC hdc = BeginPaint(window, &paint);
        DrawWindowBackground(hdc);
        EndPaint(window, &paint);
        return 0;
    }

    case WM_MEASUREITEM:
        if (wParam == kControlResults)
        {
            auto* measureItem = reinterpret_cast<MEASUREITEMSTRUCT*>(lParam);
            measureItem->itemHeight = 42;
            return TRUE;
        }
        break;

    case WM_DRAWITEM:
        if (wParam == kControlResults)
        {
            DrawResultItem(reinterpret_cast<DRAWITEMSTRUCT*>(lParam));
            return TRUE;
        }
        if (wParam == kControlAutostartToggle)
        {
            DrawAutostartToggle(reinterpret_cast<DRAWITEMSTRUCT*>(lParam));
            return TRUE;
        }
        break;

    case WM_HOTKEY:
        if (wParam == kLauncherHotkeyId)
        {
            ToggleLauncher();
            return 0;
        }
        break;

    case WM_NCHITTEST:
    {
        const LRESULT hit = DefWindowProcW(window, message, wParam, lParam);
        if (hit != HTCLIENT)
        {
            return hit;
        }

        POINT point{
            static_cast<SHORT>(LOWORD(lParam)),
            static_cast<SHORT>(HIWORD(lParam))
        };
        ScreenToClient(window, &point);

        auto isInsideChild = [&](HWND child) -> bool
        {
            if (child == nullptr || !IsWindowVisible(child))
            {
                return false;
            }

            RECT childRect{};
            GetWindowRect(child, &childRect);
            POINT topLeft{childRect.left, childRect.top};
            POINT bottomRight{childRect.right, childRect.bottom};
            ScreenToClient(window, &topLeft);
            ScreenToClient(window, &bottomRight);
            RECT localRect{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
            return PtInRect(&localRect, point) != FALSE;
        };

        if (isInsideChild(queryEdit_) || isInsideChild(resultsList_) || isInsideChild(autostartToggle_))
        {
            return HTCLIENT;
        }

        return HTCAPTION;
    }

    case WM_COMMAND:
        if (LOWORD(wParam) == kControlQuery && HIWORD(wParam) == EN_CHANGE)
        {
            ScheduleResultsUpdate();
            return 0;
        }

        if (LOWORD(wParam) == kControlResults && HIWORD(wParam) == LBN_SELCHANGE)
        {
            UpdateSelectedPath();
            return 0;
        }

        if (LOWORD(wParam) == kControlResults && HIWORD(wParam) == LBN_DBLCLK)
        {
            LaunchSelected();
            return 0;
        }

        if (LOWORD(wParam) == kControlAutostartToggle && HIWORD(wParam) == BN_CLICKED)
        {
            ToggleAutostartSetting();
            return 0;
        }

        if (LOWORD(wParam) == kCommandTrayOpen)
        {
            ShowLauncher();
            return 0;
        }

        if (LOWORD(wParam) == kCommandTrayRebuild)
        {
            StartIndexThread(true);
            return 0;
        }

        if (LOWORD(wParam) == kCommandTrayOpenSettings)
        {
            OpenSettingsFile();
            return 0;
        }

        if (LOWORD(wParam) == kCommandTrayReloadSettings)
        {
            ReloadSettingsAndRefresh();
            return 0;
        }

        if (LOWORD(wParam) == kCommandTrayToggleAutostart)
        {
            ToggleAutostartSetting();
            return 0;
        }

        if (LOWORD(wParam) == kCommandTrayExit)
        {
            DestroyWindow(mainWindow_);
            return 0;
        }

        if (LOWORD(wParam) == kCommandResultLaunch)
        {
            LaunchSelected();
            return 0;
        }

        if (LOWORD(wParam) == kCommandResultRunAsAdmin)
        {
            LaunchSelectedAsAdmin();
            return 0;
        }

        if (LOWORD(wParam) == kCommandResultOpenLocation)
        {
            OpenSelectedLocation();
            return 0;
        }

        if (LOWORD(wParam) == kCommandResultCopyPath)
        {
            CopySelectedPath();
            return 0;
        }
        break;

    case WM_CTLCOLORSTATIC:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND control = reinterpret_cast<HWND>(lParam);
        SetBkMode(hdc, TRANSPARENT);

        if (control == pathLabel_ || control == statusLabel_ || control == autostartLabel_)
        {
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, RGB(255, 255, 255));
            if (control == pathLabel_)
            {
                SetTextColor(hdc, RGB(89, 96, 107));
            }
            else if (control == statusLabel_)
            {
                SetTextColor(hdc, RGB(122, 127, 133));
            }
            else
            {
                SetTextColor(hdc, RGB(116, 120, 126));
            }
            return reinterpret_cast<LRESULT>(GetStockObject(WHITE_BRUSH));
        }

        if (control == titleLabel_)
        {
            SetTextColor(hdc, RGB(24, 24, 24));
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }

        if (control == sidebarChipLabel_)
        {
            SetTextColor(hdc, RGB(39, 52, 89));
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }

        if (control == hintLabel_)
        {
            SetTextColor(hdc, RGB(118, 121, 126));
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }

        if (control == sidebarFooterLabel_)
        {
            SetTextColor(hdc, RGB(123, 110, 96));
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }

        if (control == mainTitleLabel_)
        {
            SetTextColor(hdc, RGB(18, 18, 18));
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }

        if (control == mainSubtitleLabel_)
        {
            SetTextColor(hdc, RGB(116, 120, 126));
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }
        break;
    }

    case WM_CTLCOLOREDIT:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        SetBkColor(hdc, RGB(255, 255, 255));
        SetTextColor(hdc, RGB(28, 30, 34));
        return reinterpret_cast<LRESULT>(GetStockObject(WHITE_BRUSH));
    }

    case WM_CTLCOLORLISTBOX:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        SetBkColor(hdc, RGB(255, 255, 255));
        SetTextColor(hdc, RGB(28, 30, 34));
        return reinterpret_cast<LRESULT>(GetStockObject(WHITE_BRUSH));
    }

    case WM_TIMER:
        if (wParam == kFocusTimerId)
        {
            ReinforceFocus();
            return 0;
        }

        if (wParam == kFilesystemDebounceTimerId)
        {
            KillTimer(mainWindow_, kFilesystemDebounceTimerId);
            if (indexRunning_)
            {
                watchRefreshQueued_ = true;
            }
            else
            {
                StartIndexThread(false);
            }
            return 0;
        }

        if (wParam == kQueryDebounceTimerId)
        {
            RequestSearchUpdate();
            return 0;
        }
        break;

    case WM_ACTIVATE:
        if (LOWORD(wParam) == WA_INACTIVE && IsWindowVisible(mainWindow_))
        {
            HideLauncher();
        }
        return 0;

    case kToggleLauncherMessage:
        ToggleLauncher();
        return 0;

    case kIndexUpdatedMessage:
        queryUpdatePending_ = false;
        KillTimer(mainWindow_, kQueryDebounceTimerId);
        RequestSearchUpdate();
        UpdateFooter();
        if (watchRefreshQueued_.exchange(false))
        {
            StartIndexThread(false);
        }
        return 0;

    case kSearchResultsReadyMessage:
        ApplyReadySearchResults();
        return 0;

    case kStatusUpdatedMessage:
        UpdateFooter();
        return 0;

    case kFilesystemChangedMessage:
        SetTimer(mainWindow_, kFilesystemDebounceTimerId, 1500, nullptr);
        if (indexRunning_)
        {
            watchRefreshQueued_ = true;
        }
        return 0;

    case kTrayMessage:
        if (lParam == WM_RBUTTONUP || lParam == WM_CONTEXTMENU)
        {
            ShowTrayMenu();
        }
        else if (lParam == WM_LBUTTONDBLCLK)
        {
            ShowLauncher();
        }
        return 0;

    case WM_CLOSE:
        HideLauncher();
        return 0;

    case WM_DESTROY:
        shuttingDown_ = true;
        KillTimer(mainWindow_, kFocusTimerId);
        KillTimer(mainWindow_, kFilesystemDebounceTimerId);
        KillTimer(mainWindow_, kQueryDebounceTimerId);
        StopSearchThread();
        StopWatchThread();
        UnregisterLauncherHotkey();
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(window, message, wParam, lParam);
}

LRESULT LauncherApp::HandleEditMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
    if (message == WM_KEYDOWN)
    {
        const bool controlPressed = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

        if (controlPressed && (wParam == 'A' || wParam == 'a'))
        {
            SendMessageW(queryEdit_, EM_SETSEL, 0, -1);
            return 0;
        }

        if (wParam == VK_DOWN)
        {
            MoveSelection(1);
            return 0;
        }

        if (wParam == VK_UP)
        {
            MoveSelection(-1);
            return 0;
        }

        if (wParam == VK_RETURN)
        {
            if (controlPressed && shiftPressed)
            {
                LaunchSelectedAsAdmin();
            }
            else if (shiftPressed)
            {
                OpenSelectedLocation();
            }
            else
            {
                LaunchSelected();
            }
            return 0;
        }

        if (wParam == VK_ESCAPE)
        {
            HideLauncher();
            return 0;
        }

        if (controlPressed && (wParam == 'C' || wParam == 'c'))
        {
            DWORD selectionStart = 0;
            DWORD selectionEnd = 0;
            SendMessageW(queryEdit_, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
            if (selectionStart == selectionEnd)
            {
                CopySelectedPath();
                return 0;
            }
        }

        if (controlPressed && (wParam == 'R' || wParam == 'r'))
        {
            StartIndexThread(true);
            return 0;
        }
    }

    return DefSubclassProc(window, message, wParam, lParam);
}

LRESULT LauncherApp::HandleListMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
    if (message == WM_CONTEXTMENU)
    {
        POINT point{
            static_cast<SHORT>(LOWORD(lParam)),
            static_cast<SHORT>(HIWORD(lParam))
        };
        ShowSelectedContextMenu(point, false);
        return 0;
    }

    if (message == WM_KEYDOWN)
    {
        const bool controlPressed = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

        if (wParam == VK_RETURN)
        {
            if (controlPressed && shiftPressed)
            {
                LaunchSelectedAsAdmin();
            }
            else if (shiftPressed)
            {
                OpenSelectedLocation();
            }
            else
            {
                LaunchSelected();
            }
            return 0;
        }

        if (wParam == VK_ESCAPE)
        {
            HideLauncher();
            return 0;
        }

        if (wParam == VK_APPS || (shiftPressed && wParam == VK_F10))
        {
            POINT point{};
            ShowSelectedContextMenu(point, true);
            return 0;
        }

        if (controlPressed && (wParam == 'C' || wParam == 'c'))
        {
            CopySelectedPath();
            return 0;
        }

        if (controlPressed && (wParam == 'R' || wParam == 'r'))
        {
            StartIndexThread(true);
            return 0;
        }
    }

    return DefSubclassProc(window, message, wParam, lParam);
}

void LauncherApp::CreateFonts()
{
    titleFont_ = CreateFontW(-28, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    normalFont_ = CreateFontW(-17, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    smallFont_ = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
}

void LauncherApp::CreateControls()
{
    titleLabel_ = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    hintLabel_ = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    sidebarChipLabel_ = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    sidebarFooterLabel_ = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    mainTitleLabel_ = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    mainSubtitleLabel_ = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    queryEdit_ = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0, 0, 0, mainWindow_, reinterpret_cast<HMENU>(kControlQuery), instance_, nullptr);
    resultsList_ = CreateWindowExW(0, L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | LBS_NOTIFY | LBS_OWNERDRAWFIXED | WS_VSCROLL | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, mainWindow_, reinterpret_cast<HMENU>(kControlResults), instance_, nullptr);
    pathLabel_ = CreateWindowExW(0, L"STATIC", L"Waiting for index...", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, mainWindow_, reinterpret_cast<HMENU>(kControlPath), instance_, nullptr);
    statusLabel_ = CreateWindowExW(0, L"STATIC", L"Initializing...", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, mainWindow_, reinterpret_cast<HMENU>(kControlStatus), instance_, nullptr);
    autostartLabel_ = CreateWindowExW(0, L"STATIC", L"Run at Startup", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, mainWindow_, nullptr, instance_, nullptr);
    autostartToggle_ = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 0, 0, mainWindow_, reinterpret_cast<HMENU>(kControlAutostartToggle), instance_, nullptr);

    SendMessageW(titleLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(titleFont_), TRUE);
    SendMessageW(hintLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(smallFont_), TRUE);
    SendMessageW(sidebarChipLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(normalFont_), TRUE);
    SendMessageW(sidebarFooterLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(smallFont_), TRUE);
    SendMessageW(mainTitleLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(titleFont_), TRUE);
    SendMessageW(mainSubtitleLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(smallFont_), TRUE);
    SendMessageW(queryEdit_, WM_SETFONT, reinterpret_cast<WPARAM>(normalFont_), TRUE);
    SendMessageW(resultsList_, WM_SETFONT, reinterpret_cast<WPARAM>(normalFont_), TRUE);
    SendMessageW(pathLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(smallFont_), TRUE);
    SendMessageW(statusLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(smallFont_), TRUE);
    SendMessageW(autostartLabel_, WM_SETFONT, reinterpret_cast<WPARAM>(smallFont_), TRUE);
    SendMessageW(queryEdit_, EM_SETCUEBANNER, TRUE, reinterpret_cast<LPARAM>(L"Search app, script, folder or path"));
    SendMessageW(queryEdit_, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(10, 10));
    SendMessageW(resultsList_, LB_SETITEMHEIGHT, 0, 42);

    SetWindowSubclass(queryEdit_, EditSubclassProc, 1, reinterpret_cast<DWORD_PTR>(this));
    SetWindowSubclass(resultsList_, ListSubclassProc, 2, reinterpret_cast<DWORD_PTR>(this));
    EnsureShellImageList();
    UpdateHintLabel();
    RefreshAutostartToggle();
}

void LauncherApp::LayoutControls()
{
    RECT rect{};
    GetClientRect(mainWindow_, &rect);

    const int width = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    const int contentLeft = 14;
    const int contentWidth = width - 28;
    const int searchCardTop = 14;
    const int searchCardWidth = contentWidth;
    const int searchCardLeft = contentLeft;
    const int searchCardHeight = 42;
    const int resultsCardTop = 64;
    const int resultsCardHeight = 286;
    const int resultsListHeight = 42 * 5;
    const int toggleWidth = 54;
    const int toggleHeight = 26;
    const int queryHeight = 20;

    HRGN region = CreateRoundRectRgn(0, 0, width + 1, height + 1, 26, 26);
    SetWindowRgn(mainWindow_, region, TRUE);

    MoveWindow(titleLabel_, 0, 0, 0, 0, FALSE);
    MoveWindow(hintLabel_, 0, 0, 0, 0, FALSE);
    MoveWindow(sidebarChipLabel_, 0, 0, 0, 0, FALSE);
    MoveWindow(sidebarFooterLabel_, 0, 0, 0, 0, FALSE);
    MoveWindow(mainTitleLabel_, 0, 0, 0, 0, FALSE);
    MoveWindow(mainSubtitleLabel_, 0, 0, 0, 0, FALSE);

    MoveWindow(queryEdit_, searchCardLeft + 12, searchCardTop + (searchCardHeight - queryHeight) / 2, searchCardWidth - 24, queryHeight, TRUE);
    MoveWindow(resultsList_, contentLeft + 12, resultsCardTop + 12, contentWidth - 24, resultsListHeight, TRUE);
    MoveWindow(pathLabel_, contentLeft + 16, resultsCardTop + 236, contentWidth - 170, 18, TRUE);
    MoveWindow(statusLabel_, contentLeft + 16, resultsCardTop + 258, contentWidth - 170, 18, TRUE);
    MoveWindow(autostartLabel_, contentLeft + contentWidth - 128, resultsCardTop + 238, 68, 18, TRUE);
    MoveWindow(autostartToggle_, contentLeft + contentWidth - toggleWidth - 16, resultsCardTop + 232, toggleWidth, toggleHeight, TRUE);
}

void LauncherApp::CreateTrayIcon()
{
    ZeroMemory(&trayIconData_, sizeof(trayIconData_));
    trayIconData_.cbSize = sizeof(trayIconData_);
    trayIconData_.hWnd = mainWindow_;
    trayIconData_.uID = 1;
    trayIconData_.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    trayIconData_.uCallbackMessage = kTrayMessage;
    trayIconData_.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
    wcscpy_s(trayIconData_.szTip, L"QuickLauncher Native");
    Shell_NotifyIconW(NIM_ADD, &trayIconData_);
}

void LauncherApp::ShowTrayMenu()
{
    autostartEnabled_ = IsAutostartEnabled();
    RefreshAutostartToggle();

    HMENU menu = CreatePopupMenu();
    AppendMenuW(menu, MF_STRING, kCommandTrayOpen, L"Open Launcher");
    AppendMenuW(menu, MF_STRING, kCommandTrayRebuild, L"Rebuild Index");
    AppendMenuW(menu, MF_STRING, kCommandTrayReloadSettings, L"Reload Settings");
    AppendMenuW(menu, MF_STRING, kCommandTrayOpenSettings, L"Open Settings");
    AppendMenuW(menu, autostartEnabled_ ? MF_STRING | MF_CHECKED : MF_STRING, kCommandTrayToggleAutostart, L"Run at Startup");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, kCommandTrayExit, L"Exit");

    POINT point{};
    GetCursorPos(&point);
    SetForegroundWindow(mainWindow_);
    TrackPopupMenu(menu, TPM_RIGHTBUTTON, point.x, point.y, 0, mainWindow_, nullptr);
    DestroyMenu(menu);
}

void LauncherApp::EnsureSettingsFileExists() const
{
    std::ifstream existing(Utf8FromWide(settingsPath_), std::ios::binary);
    if (existing.is_open())
    {
        return;
    }

    std::ofstream stream(Utf8FromWide(settingsPath_), std::ios::binary | std::ios::trunc);
    if (!stream.is_open())
    {
        return;
    }

    WriteUtf8Line(stream, L"# QuickLauncher Native settings");
    WriteUtf8Line(stream, L"# Use one item per line.");
    WriteUtf8Line(stream, L"# hotkey=Alt+Space");
    WriteUtf8Line(stream, L"# include=E:\\Tools");
    WriteUtf8Line(stream, L"# exclude=E:\\SteamLibrary\\steamapps");
    WriteUtf8Line(stream, L"# priority=E:\\Tools");
    WriteUtf8Line(stream, L"");
}

void LauncherApp::LoadSettings()
{
    std::ifstream stream(Utf8FromWide(settingsPath_), std::ios::binary);
    if (!stream.is_open())
    {
        PublishStatus(L"Failed to open settings.");
        return;
    }

    LauncherSettings loaded;
    std::wstring invalidHotkeyValue;
    std::string utf8Line;

    while (std::getline(stream, utf8Line))
    {
        if (!utf8Line.empty() && utf8Line.back() == '\r')
        {
            utf8Line.pop_back();
        }

        auto line = TrimCopy(WideFromUtf8(utf8Line));
        if (line.empty() || line[0] == L'#' || line[0] == L';')
        {
            continue;
        }

        const auto separator = line.find(L'=');
        if (separator == std::wstring::npos)
        {
            continue;
        }

        auto key = ToLowerCopy(TrimCopy(line.substr(0, separator)));
        auto rawValue = TrimCopy(line.substr(separator + 1));
        if (key == L"hotkey")
        {
            HotkeyParseResult parsed;
            if (TryParseHotkey(rawValue, parsed))
            {
                loaded.hotkeyModifiers = parsed.modifiers;
                loaded.hotkeyVirtualKey = parsed.virtualKey;
                loaded.hotkeyLabel = parsed.label;
            }
            else if (invalidHotkeyValue.empty())
            {
                invalidHotkeyValue = rawValue;
            }

            continue;
        }

        auto value = NormalizeDirectoryPath(rawValue);
        if (value.empty())
        {
            continue;
        }

        if (key == L"include")
        {
            loaded.includeDirectories.push_back(value);
        }
        else if (key == L"exclude")
        {
            loaded.excludeDirectories.push_back(value);
        }
        else if (key == L"priority")
        {
            loaded.priorityDirectories.push_back(value);
        }
    }

    auto dedupe = [](std::vector<std::wstring>& values)
    {
        std::sort(values.begin(), values.end());
        values.erase(std::unique(values.begin(), values.end()), values.end());
    };

    dedupe(loaded.includeDirectories);
    dedupe(loaded.excludeDirectories);
    dedupe(loaded.priorityDirectories);

    {
        ScopedCriticalSection lock(dataLock_);
        settings_ = std::move(loaded);
    }

    if (invalidHotkeyValue.empty())
    {
        PublishStatus(L"Settings loaded.");
    }
    else
    {
        PublishStatus(L"Invalid hotkey setting ignored. Using Alt+Space.");
    }
}

void LauncherApp::OpenSettingsFile() const
{
    EnsureSettingsFileExists();
    ShellExecuteW(mainWindow_, L"open", settingsPath_.c_str(), nullptr, dataDirectory_.c_str(), SW_SHOWNORMAL);
}

void LauncherApp::ReloadSettingsAndRefresh()
{
    LoadSettings();
    RegisterLauncherHotkey();
    RestartWatchThread();
    RebuildSearchCaches();
    UpdateResults();
    StartIndexThread(true);
}

bool LauncherApp::IsAutostartEnabled() const
{
    HKEY key = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_READ, &key) != ERROR_SUCCESS)
    {
        return false;
    }

    wchar_t buffer[2048] = {};
    DWORD type = 0;
    DWORD size = sizeof(buffer);
    const auto result = RegQueryValueExW(key, L"QuickLauncherNative", nullptr, &type, reinterpret_cast<LPBYTE>(buffer), &size);
    RegCloseKey(key);

    if (result != ERROR_SUCCESS || type != REG_SZ)
    {
        return false;
    }

    return ToLowerCopy(buffer).find(ToLowerCopy(GetExecutablePath())) != std::wstring::npos;
}

bool LauncherApp::SetAutostartEnabled(bool enabled) const
{
    HKEY key = nullptr;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, nullptr, 0, KEY_SET_VALUE, nullptr, &key, nullptr) != ERROR_SUCCESS)
    {
        return false;
    }

    bool success = true;
    if (enabled)
    {
        const auto value = L"\"" + GetExecutablePath() + L"\"";
        const DWORD size = static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t));
        success = RegSetValueExW(key, L"QuickLauncherNative", 0, REG_SZ, reinterpret_cast<const BYTE*>(value.c_str()), size) == ERROR_SUCCESS;
    }
    else
    {
        const auto result = RegDeleteValueW(key, L"QuickLauncherNative");
        success = result == ERROR_SUCCESS || result == ERROR_FILE_NOT_FOUND;
    }

    RegCloseKey(key);
    return success;
}

void LauncherApp::RegisterLauncherHotkey()
{
    UnregisterLauncherHotkey();

    UINT modifiers = MOD_ALT;
    UINT virtualKey = VK_SPACE;
    std::wstring hotkeyLabel = L"Alt+Space";
    {
        ScopedCriticalSection lock(dataLock_);
        modifiers = settings_.hotkeyModifiers;
        virtualKey = settings_.hotkeyVirtualKey;
        hotkeyLabel = settings_.hotkeyLabel;
    }

    if (mainWindow_ == nullptr)
    {
        return;
    }

    if (RegisterHotKey(mainWindow_, kLauncherHotkeyId, modifiers | MOD_NOREPEAT, virtualKey))
    {
        activeHotkeyLabel_ = hotkeyLabel;
        UpdateHintLabel();
        PublishStatus(L"Hotkey active: " + activeHotkeyLabel_ + L".");
        return;
    }

    if ((modifiers != MOD_ALT || virtualKey != VK_SPACE)
        && RegisterHotKey(mainWindow_, kLauncherHotkeyId, MOD_ALT | MOD_NOREPEAT, VK_SPACE))
    {
        activeHotkeyLabel_ = L"Alt+Space";
        UpdateHintLabel();
        PublishStatus(L"Requested hotkey unavailable. Using Alt+Space.");
        return;
    }

    activeHotkeyLabel_.clear();
    UpdateHintLabel();
    PublishStatus(L"Failed to register global hotkey. Open from tray.");
}

void LauncherApp::UnregisterLauncherHotkey()
{
    if (mainWindow_ != nullptr)
    {
        UnregisterHotKey(mainWindow_, kLauncherHotkeyId);
    }
}

void LauncherApp::UpdateHintLabel()
{
    if (hintLabel_ != nullptr)
    {
        SetWindowTextW(hintLabel_, L"");
    }
    if (mainSubtitleLabel_ != nullptr)
    {
        SetWindowTextW(mainSubtitleLabel_, L"");
    }
    if (sidebarFooterLabel_ != nullptr)
    {
        SetWindowTextW(sidebarFooterLabel_, L"");
    }
}

void LauncherApp::RefreshAutostartToggle()
{
    if (autostartToggle_ != nullptr)
    {
        InvalidateRect(autostartToggle_, nullptr, TRUE);
        UpdateWindow(autostartToggle_);
    }
}

void LauncherApp::ToggleAutostartSetting()
{
    const bool nextEnabled = !autostartEnabled_;
    if (!SetAutostartEnabled(nextEnabled))
    {
        PublishStatus(L"Failed to update autostart.");
        UpdateFooter();
        return;
    }

    autostartEnabled_ = nextEnabled;
    RefreshAutostartToggle();
    PublishStatus(autostartEnabled_ ? L"Autostart enabled." : L"Autostart disabled.");
    UpdateFooter();
}

void LauncherApp::StartWatchThread()
{
    StopWatchThread();

    watchStopEvent_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (watchStopEvent_ == nullptr)
    {
        PublishStatus(L"Failed to start filesystem watcher.");
        return;
    }

    watchThread_ = CreateThread(nullptr, 0, WatchThreadProc, this, 0, nullptr);
    if (watchThread_ == nullptr)
    {
        CloseHandle(watchStopEvent_);
        watchStopEvent_ = nullptr;
        PublishStatus(L"Failed to start filesystem watcher.");
    }
}

void LauncherApp::StopWatchThread()
{
    if (watchStopEvent_ != nullptr)
    {
        SetEvent(watchStopEvent_);
    }

    if (watchThread_ != nullptr)
    {
        WaitForSingleObject(watchThread_, 3000);
        CloseHandle(watchThread_);
        watchThread_ = nullptr;
    }

    if (watchStopEvent_ != nullptr)
    {
        CloseHandle(watchStopEvent_);
        watchStopEvent_ = nullptr;
    }
}

void LauncherApp::RestartWatchThread()
{
    if (shuttingDown_)
    {
        return;
    }

    StartWatchThread();
}

void LauncherApp::StartSearchThread()
{
    StopSearchThread();

    searchStopEvent_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    searchWakeEvent_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (searchStopEvent_ == nullptr || searchWakeEvent_ == nullptr)
    {
        if (searchStopEvent_ != nullptr)
        {
            CloseHandle(searchStopEvent_);
            searchStopEvent_ = nullptr;
        }

        if (searchWakeEvent_ != nullptr)
        {
            CloseHandle(searchWakeEvent_);
            searchWakeEvent_ = nullptr;
        }

        PublishStatus(L"Failed to start background search.");
        return;
    }

    searchThread_ = CreateThread(nullptr, 0, SearchThreadProc, this, 0, nullptr);
    if (searchThread_ == nullptr)
    {
        CloseHandle(searchStopEvent_);
        CloseHandle(searchWakeEvent_);
        searchStopEvent_ = nullptr;
        searchWakeEvent_ = nullptr;
        PublishStatus(L"Failed to start background search.");
    }
}

void LauncherApp::StopSearchThread()
{
    if (searchStopEvent_ != nullptr)
    {
        SetEvent(searchStopEvent_);
    }

    if (searchWakeEvent_ != nullptr)
    {
        SetEvent(searchWakeEvent_);
    }

    if (searchThread_ != nullptr)
    {
        WaitForSingleObject(searchThread_, 3000);
        CloseHandle(searchThread_);
        searchThread_ = nullptr;
    }

    if (searchWakeEvent_ != nullptr)
    {
        CloseHandle(searchWakeEvent_);
        searchWakeEvent_ = nullptr;
    }

    if (searchStopEvent_ != nullptr)
    {
        CloseHandle(searchStopEvent_);
        searchStopEvent_ = nullptr;
    }
}

void LauncherApp::WatchLoop()
{
    const auto roots = BuildWatchRoots();

    std::vector<HANDLE> handles;
    handles.reserve(std::min<size_t>(roots.size() + 1, MAXIMUM_WAIT_OBJECTS));
    handles.push_back(watchStopEvent_);

    for (const auto& root : roots)
    {
        if (handles.size() >= MAXIMUM_WAIT_OBJECTS)
        {
            break;
        }

        const HANDLE changeHandle = FindFirstChangeNotificationW(root.c_str(), TRUE, kWatchNotifyFilter);
        if (changeHandle != INVALID_HANDLE_VALUE)
        {
            handles.push_back(changeHandle);
        }
    }

    if (handles.size() <= 1)
    {
        return;
    }

    PublishStatus(L"Watching quick-access folders.");

    while (!shuttingDown_)
    {
        const DWORD waitResult = WaitForMultipleObjects(static_cast<DWORD>(handles.size()), handles.data(), FALSE, INFINITE);
        if (waitResult == WAIT_OBJECT_0)
        {
            break;
        }

        if (waitResult > WAIT_OBJECT_0 && waitResult < WAIT_OBJECT_0 + handles.size())
        {
            const size_t handleIndex = static_cast<size_t>(waitResult - WAIT_OBJECT_0);
            if (mainWindow_ != nullptr)
            {
                PostMessageW(mainWindow_, kFilesystemChangedMessage, 0, 0);
            }

            if (!FindNextChangeNotification(handles[handleIndex]))
            {
                break;
            }

            continue;
        }

        break;
    }

    for (size_t index = 1; index < handles.size(); ++index)
    {
        FindCloseChangeNotification(handles[index]);
    }
}

void LauncherApp::SearchLoop()
{
    HANDLE handles[2] = {searchStopEvent_, searchWakeEvent_};

    while (!shuttingDown_)
    {
        const DWORD waitResult = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
        if (waitResult == WAIT_OBJECT_0)
        {
            break;
        }

        if (waitResult != WAIT_OBJECT_0 + 1)
        {
            break;
        }

        ResetEvent(searchWakeEvent_);

        while (!shuttingDown_)
        {
            std::wstring query;
            std::shared_ptr<std::vector<LaunchEntry>> entriesSnapshot;
            std::unordered_map<std::wstring, UsageStat> usageSnapshot;
            std::unordered_map<std::wstring, QueryBucket> queryHistorySnapshot;
            std::vector<std::wstring> excludedRoots;
            std::vector<std::wstring> priorityRoots;
            bool hasDirectEntry = false;
            LaunchEntry directEntry;
            std::vector<size_t> candidateIndices;
            unsigned long long requestId = 0;

            {
                ScopedCriticalSection lock(dataLock_);
                if (!searchRequestPending_)
                {
                    break;
                }

                searchRequestPending_ = false;
                query = pendingSearchQuery_;
                entriesSnapshot = pendingSearchEntries_;
                usageSnapshot = pendingSearchUsage_;
                queryHistorySnapshot = pendingSearchHistory_;
                excludedRoots = pendingSearchExcludedRoots_;
                priorityRoots = pendingSearchPriorityRoots_;
                hasDirectEntry = pendingSearchHasDirectEntry_;
                directEntry = pendingSearchDirectEntry_;
                candidateIndices = pendingSearchCandidateIndices_;
                requestId = pendingSearchRequestId_;
            }

            std::vector<size_t> matchedIndices;
            auto results = BuildResultsForQuery(
                query,
                entriesSnapshot,
                usageSnapshot,
                queryHistorySnapshot,
                excludedRoots,
                priorityRoots,
                requestId,
                candidateIndices.empty() ? nullptr : &candidateIndices,
                &matchedIndices);

            if (requestId != latestSearchRequestId_.load())
            {
                continue;
            }

            {
                ScopedCriticalSection lock(dataLock_);
                readySearchRequestId_ = requestId;
                readySearchEntries_ = entriesSnapshot;
                readySearchResults_ = std::move(results);
                readySearchHasDirectEntry_ = hasDirectEntry;
                readySearchDirectEntry_ = directEntry;
                readySearchQueryCompact_ = NormalizeCompact(query);
                readySearchMatchedIndices_ = std::move(matchedIndices);
            }

            if (mainWindow_ != nullptr)
            {
                PostMessageW(mainWindow_, kSearchResultsReadyMessage, 0, 0);
            }
        }
    }
}

void LauncherApp::InitializePreferredRoots()
{
    desktopRoots_.clear();
    startMenuRoots_.clear();
    pinnedRoots_.clear();

    auto appendKnownFolder = [](std::vector<std::wstring>& roots, int csidl)
    {
        wchar_t buffer[MAX_PATH * 4] = {};
        if (SUCCEEDED(SHGetFolderPathW(nullptr, csidl, nullptr, SHGFP_TYPE_CURRENT, buffer)))
        {
            AppendUniqueRoot(roots, buffer);
        }
    };

    appendKnownFolder(desktopRoots_, CSIDL_DESKTOPDIRECTORY);
    appendKnownFolder(desktopRoots_, CSIDL_COMMON_DESKTOPDIRECTORY);
    appendKnownFolder(startMenuRoots_, CSIDL_PROGRAMS);
    appendKnownFolder(startMenuRoots_, CSIDL_COMMON_PROGRAMS);

    wchar_t appDataBuffer[MAX_PATH * 4] = {};
    if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, SHGFP_TYPE_CURRENT, appDataBuffer)))
    {
        AppendUniqueRoot(pinnedRoots_, std::wstring(appDataBuffer) + L"\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\StartMenu");
        AppendUniqueRoot(pinnedRoots_, std::wstring(appDataBuffer) + L"\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar");
    }
}

void LauncherApp::ShowLauncher()
{
    autostartEnabled_ = IsAutostartEnabled();
    RefreshAutostartToggle();
    SetWindowTextW(queryEdit_, L"");
    queryUpdatePending_ = false;
    KillTimer(mainWindow_, kQueryDebounceTimerId);
    UpdateResults();

    PositionWindow();
    ShowWindow(mainWindow_, SW_SHOW);
    SetWindowPos(mainWindow_, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    BringWindowToTop(mainWindow_);
    SetForegroundWindow(mainWindow_);
    focusRetryCount_ = 10;
    SetTimer(mainWindow_, kFocusTimerId, 60, nullptr);
    FocusQueryBox(false);
}

void LauncherApp::HideLauncher()
{
    KillTimer(mainWindow_, kFocusTimerId);
    ShowWindow(mainWindow_, SW_HIDE);
}

void LauncherApp::ToggleLauncher()
{
    if (IsWindowVisible(mainWindow_))
    {
        HideLauncher();
        return;
    }

    ShowLauncher();
}

void LauncherApp::ReinforceFocus()
{
    if (!IsWindowVisible(mainWindow_))
    {
        KillTimer(mainWindow_, kFocusTimerId);
        return;
    }

    if (focusRetryCount_-- <= 0)
    {
        KillTimer(mainWindow_, kFocusTimerId);
        return;
    }

    if (GetForegroundWindow() != mainWindow_)
    {
        SetForegroundWindow(mainWindow_);
        BringWindowToTop(mainWindow_);
    }

    FocusQueryBox(false);
}

void LauncherApp::FocusQueryBox(bool selectAll)
{
    SetFocus(queryEdit_);

    const int textLength = GetWindowTextLengthW(queryEdit_);
    if (selectAll)
    {
        SendMessageW(queryEdit_, EM_SETSEL, 0, textLength);
    }
    else
    {
        SendMessageW(queryEdit_, EM_SETSEL, textLength, textLength);
    }
}

void LauncherApp::PositionWindow()
{
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize = sizeof(monitorInfo);
    GetMonitorInfoW(MonitorFromWindow(mainWindow_, MONITOR_DEFAULTTOPRIMARY), &monitorInfo);

    const RECT& work = monitorInfo.rcWork;
    const int width = 880;
    const int height = 382;
    const int x = work.left + ((work.right - work.left) - width) / 2;
    const int y = work.top + std::max(26, static_cast<int>((work.bottom - work.top) / 10));

    SetWindowPos(mainWindow_, HWND_TOPMOST, x, y, width, height, SWP_SHOWWINDOW);
}

void LauncherApp::RebuildSearchCaches()
{
    std::shared_ptr<std::vector<LaunchEntry>> entriesSnapshot;
    std::unordered_map<std::wstring, UsageStat> usageSnapshot;
    std::vector<std::wstring> priorityRoots;

    {
        ScopedCriticalSection lock(dataLock_);
        entriesSnapshot = entries_;
        usageSnapshot = usage_;
        priorityRoots = settings_.priorityDirectories;
    }

    std::unordered_map<std::wstring, std::vector<size_t>> prefixIndex;
    if (entriesSnapshot != nullptr)
    {
        prefixIndex.reserve(1024);
        const auto& entries = *entriesSnapshot;
        for (size_t index = 0; index < entries.size(); ++index)
        {
            std::unordered_set<std::wstring> keys;
            AppendPrefixKeys(keys, entries[index].searchNameCompact);
            for (const auto& token : TokenizeQuery(entries[index].name))
            {
                AppendPrefixKeys(keys, token);
            }

            AppendPrefixKeys(keys, entries[index].searchInitials);
            for (const auto& key : keys)
            {
                prefixIndex[key].push_back(index);
            }
        }
    }

    auto emptyQueryCandidates = BuildEmptyQueryCandidateIndices(entriesSnapshot, usageSnapshot, priorityRoots);

    {
        ScopedCriticalSection lock(dataLock_);
        if (entries_ != entriesSnapshot)
        {
            return;
        }

        searchCacheEntries_ = entriesSnapshot;
        searchPrefixIndex_ = std::move(prefixIndex);
        emptyQueryCandidateIndices_ = std::move(emptyQueryCandidates);
        readySearchQueryCompact_.clear();
        readySearchMatchedIndices_.clear();
    }
}

void LauncherApp::UpdateResults()
{
    const auto query = ReadControlText(queryEdit_);
    const auto queryCompact = NormalizeCompact(query);
    const auto queryTokens = TokenizeQuery(query);
    LaunchEntry directEntry;
    const bool hasDirectEntry = TryBuildDirectEntry(query, directEntry);
    std::shared_ptr<std::vector<LaunchEntry>> entriesSnapshot;
    std::unordered_map<std::wstring, UsageStat> usageSnapshot;
    std::unordered_map<std::wstring, QueryBucket> queryHistorySnapshot;
    std::vector<std::wstring> excludedRoots;
    std::vector<std::wstring> priorityRoots;
    std::vector<size_t> candidateIndices;

    {
        ScopedCriticalSection lock(dataLock_);
        entriesSnapshot = entries_;
        usageSnapshot = usage_;
        queryHistorySnapshot = queryHistory_;
        excludedRoots = settings_.excludeDirectories;
        priorityRoots = settings_.priorityDirectories;
        candidateIndices = BuildCandidateIndicesLocked(query, queryCompact, queryTokens, entriesSnapshot);
    }

    auto results = BuildResultsForQuery(
        query,
        entriesSnapshot,
        usageSnapshot,
        queryHistorySnapshot,
        excludedRoots,
        priorityRoots,
        0,
        candidateIndices.empty() ? nullptr : &candidateIndices,
        nullptr);

    currentEntries_ = entriesSnapshot;
    currentResults_ = std::move(results);
    hasDirectEntry_ = hasDirectEntry;
    directEntry_ = directEntry;
    if (hasDirectEntry_)
    {
        const auto directPathKey = NormalizeDirectoryPath(directEntry_.fullPath);
        currentResults_.erase(
            std::remove_if(currentResults_.begin(), currentResults_.end(), [&](const SearchResult& result)
            {
                return result.entryIndex < currentEntries_->size()
                    && NormalizeDirectoryPath((*currentEntries_)[result.entryIndex].fullPath) == directPathKey;
            }),
            currentResults_.end());

        currentResults_.insert(currentResults_.begin(), SearchResult{SIZE_MAX, 100000, 0});
        if (currentResults_.size() > 32)
        {
            currentResults_.pop_back();
        }
    }
    PopulateResultsList();
    UpdateFooter();
}

void LauncherApp::ScheduleResultsUpdate()
{
    queryUpdatePending_ = true;
    SetTimer(mainWindow_, kQueryDebounceTimerId, 120, nullptr);
}

void LauncherApp::RequestSearchUpdate()
{
    queryUpdatePending_ = false;
    KillTimer(mainWindow_, kQueryDebounceTimerId);

    if (searchWakeEvent_ == nullptr)
    {
        UpdateResults();
        return;
    }

    const auto query = ReadControlText(queryEdit_);
    const auto queryCompact = NormalizeCompact(query);
    const auto queryTokens = TokenizeQuery(query);
    LaunchEntry directEntry;
    const bool hasDirectEntry = TryBuildDirectEntry(query, directEntry);

    {
        ScopedCriticalSection lock(dataLock_);
        pendingSearchQuery_ = query;
        pendingSearchEntries_ = entries_;
        pendingSearchUsage_ = usage_;
        pendingSearchHistory_ = queryHistory_;
        pendingSearchExcludedRoots_ = settings_.excludeDirectories;
        pendingSearchPriorityRoots_ = settings_.priorityDirectories;
        pendingSearchHasDirectEntry_ = hasDirectEntry;
        pendingSearchDirectEntry_ = directEntry;
        pendingSearchCandidateIndices_ = BuildCandidateIndicesLocked(query, queryCompact, queryTokens, pendingSearchEntries_);
        pendingSearchRequestId_ = ++latestSearchRequestId_;
        searchRequestPending_ = true;
    }

    SetEvent(searchWakeEvent_);
}

void LauncherApp::FlushPendingResultsUpdate()
{
    bool needsSynchronousUpdate = queryUpdatePending_;
    {
        ScopedCriticalSection lock(dataLock_);
        if (!needsSynchronousUpdate && latestSearchRequestId_.load() != readySearchRequestId_)
        {
            needsSynchronousUpdate = true;
        }
    }

    if (needsSynchronousUpdate)
    {
        queryUpdatePending_ = false;
        KillTimer(mainWindow_, kQueryDebounceTimerId);
        UpdateResults();
    }
}

void LauncherApp::ApplyReadySearchResults()
{
    std::shared_ptr<std::vector<LaunchEntry>> readyEntries;
    std::vector<SearchResult> readyResults;
    LaunchEntry readyDirectEntry;
    bool readyHasDirectEntry = false;
    unsigned long long readyRequestId = 0;

    {
        ScopedCriticalSection lock(dataLock_);
        readyRequestId = readySearchRequestId_;
        readyEntries = readySearchEntries_;
        readyResults = readySearchResults_;
        readyDirectEntry = readySearchDirectEntry_;
        readyHasDirectEntry = readySearchHasDirectEntry_;
    }

    if (readyEntries == nullptr || readyRequestId != latestSearchRequestId_.load())
    {
        return;
    }

    currentEntries_ = readyEntries;
    currentResults_ = std::move(readyResults);
    hasDirectEntry_ = readyHasDirectEntry;
    directEntry_ = readyDirectEntry;
    if (hasDirectEntry_)
    {
        const auto directPathKey = NormalizeDirectoryPath(directEntry_.fullPath);
        currentResults_.erase(
            std::remove_if(currentResults_.begin(), currentResults_.end(), [&](const SearchResult& result)
            {
                return result.entryIndex < currentEntries_->size()
                    && NormalizeDirectoryPath((*currentEntries_)[result.entryIndex].fullPath) == directPathKey;
            }),
            currentResults_.end());

        currentResults_.insert(currentResults_.begin(), SearchResult{SIZE_MAX, 100000, 0});
        if (currentResults_.size() > 32)
        {
            currentResults_.pop_back();
        }
    }

    PopulateResultsList();
    UpdateFooter();
}

void LauncherApp::PopulateResultsList()
{
    SendMessageW(resultsList_, WM_SETREDRAW, FALSE, 0);
    SendMessageW(resultsList_, LB_RESETCONTENT, 0, 0);

    ScopedCriticalSection lock(dataLock_);
    for (const auto& result : currentResults_)
    {
        const auto& entry = result.entryIndex == SIZE_MAX ? directEntry_ : (*currentEntries_)[result.entryIndex];
        const std::wstring label = BuildDisplayLabel(entry);
        SendMessageW(resultsList_, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));
    }

    if (!currentResults_.empty())
    {
        SendMessageW(resultsList_, LB_SETCURSEL, 0, 0);
    }

    SendMessageW(resultsList_, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(resultsList_, nullptr, TRUE);
    UpdateSelectedPath();
}

void LauncherApp::UpdateSelectedPath()
{
    const auto selectedIndex = static_cast<int>(SendMessageW(resultsList_, LB_GETCURSEL, 0, 0));

    if (selectedIndex >= 0 && static_cast<size_t>(selectedIndex) < currentResults_.size())
    {
        if (currentResults_[selectedIndex].entryIndex == SIZE_MAX)
        {
            SetWindowTextW(pathLabel_, directEntry_.fullPath.c_str());
            return;
        }

        ScopedCriticalSection lock(dataLock_);
        SetWindowTextW(pathLabel_, (*currentEntries_)[currentResults_[selectedIndex].entryIndex].fullPath.c_str());
        return;
    }

    SetWindowTextW(pathLabel_, currentResults_.empty()
        ? L"No match. Keep typing or press Ctrl+R to rebuild."
        : L"Use Up/Down to move, Enter to launch.");
}

void LauncherApp::UpdateFooter()
{
    size_t entryCount = 0;
    std::wstring status;
    long long indexedTicks = 0;

    {
        ScopedCriticalSection lock(dataLock_);
        entryCount = entries_ == nullptr ? 0 : entries_->size();
        status = statusText_;
        indexedTicks = lastIndexedUtcTicks_;
    }

    std::wstringstream stream;
    stream << L"Indexed " << entryCount << L" items | Updated " << TicksToLocalTimeString(indexedTicks) << L" | " << status;
    SetWindowTextW(statusLabel_, stream.str().c_str());
}

void LauncherApp::MoveSelection(int offset)
{
    FlushPendingResultsUpdate();

    const int count = static_cast<int>(SendMessageW(resultsList_, LB_GETCOUNT, 0, 0));
    if (count <= 0)
    {
        return;
    }

    int current = static_cast<int>(SendMessageW(resultsList_, LB_GETCURSEL, 0, 0));
    if (current < 0)
    {
        current = 0;
    }

    current = std::clamp(current + offset, 0, count - 1);
    SendMessageW(resultsList_, LB_SETCURSEL, current, 0);
    UpdateSelectedPath();
}

void LauncherApp::LaunchSelected()
{
    FlushPendingResultsUpdate();

    LaunchEntry entry;
    if (!TryGetSelectedEntry(entry))
    {
        return;
    }

    const auto query = ReadControlText(queryEdit_);

    if (TryLaunch(entry, query))
    {
        HideLauncher();
    }
}

void LauncherApp::LaunchSelectedAsAdmin()
{
    FlushPendingResultsUpdate();

    LaunchEntry entry;
    if (!TryGetSelectedEntry(entry))
    {
        return;
    }

    const auto query = ReadControlText(queryEdit_);
    if (TryLaunchAsAdmin(entry, query))
    {
        HideLauncher();
    }
}

void LauncherApp::OpenSelectedLocation()
{
    FlushPendingResultsUpdate();

    LaunchEntry entry;
    if (!TryGetSelectedEntry(entry))
    {
        return;
    }

    if (OpenEntryLocation(entry))
    {
        HideLauncher();
    }
}

void LauncherApp::CopySelectedPath()
{
    FlushPendingResultsUpdate();

    LaunchEntry entry;
    if (!TryGetSelectedEntry(entry))
    {
        PublishStatus(L"No result selected.");
        return;
    }

    if (CopyTextToClipboard(entry.fullPath))
    {
        PublishStatus(L"Copied path to clipboard.");
    }
    else
    {
        PublishStatus(L"Failed to copy path.");
    }
}

void LauncherApp::ShowSelectedContextMenu(POINT screenPoint, bool fromKeyboard)
{
    FlushPendingResultsUpdate();

    const int count = static_cast<int>(SendMessageW(resultsList_, LB_GETCOUNT, 0, 0));
    if (count <= 0)
    {
        return;
    }

    if (!fromKeyboard)
    {
        if (screenPoint.x == -1 && screenPoint.y == -1)
        {
            fromKeyboard = true;
        }
        else
        {
            POINT clientPoint = screenPoint;
            ScreenToClient(resultsList_, &clientPoint);
            const DWORD itemFromPoint = static_cast<DWORD>(SendMessageW(
                resultsList_,
                LB_ITEMFROMPOINT,
                0,
                MAKELPARAM(clientPoint.x, clientPoint.y)));
            if (HIWORD(itemFromPoint) == 0)
            {
                const int itemIndex = static_cast<int>(LOWORD(itemFromPoint));
                SendMessageW(resultsList_, LB_SETCURSEL, itemIndex, 0);
                UpdateSelectedPath();
            }
        }
    }

    LaunchEntry entry;
    if (!TryGetSelectedEntry(entry))
    {
        return;
    }

    if (fromKeyboard)
    {
        const int selectedIndex = static_cast<int>(SendMessageW(resultsList_, LB_GETCURSEL, 0, 0));
        RECT itemRect{};
        if (selectedIndex >= 0 && SendMessageW(resultsList_, LB_GETITEMRECT, selectedIndex, reinterpret_cast<LPARAM>(&itemRect)) != LB_ERR)
        {
            POINT anchor{itemRect.left + 24, itemRect.top + (itemRect.bottom - itemRect.top) / 2};
            ClientToScreen(resultsList_, &anchor);
            screenPoint = anchor;
        }
        else
        {
            GetCursorPos(&screenPoint);
        }
    }

    HMENU menu = CreatePopupMenu();
    if (menu == nullptr)
    {
        return;
    }

    AppendMenuW(menu, MF_STRING, kCommandResultLaunch, L"Launch");
    AppendMenuW(menu, CanRunAsAdmin(entry) ? MF_STRING : MF_STRING | MF_GRAYED, kCommandResultRunAsAdmin, L"Run as Administrator");
    AppendMenuW(menu, MF_STRING, kCommandResultOpenLocation, entry.isDirectory ? L"Open Folder" : L"Open File Location");
    AppendMenuW(menu, MF_STRING, kCommandResultCopyPath, L"Copy Full Path");

    SetForegroundWindow(mainWindow_);
    TrackPopupMenu(menu, TPM_RIGHTBUTTON, screenPoint.x, screenPoint.y, 0, mainWindow_, nullptr);
    DestroyMenu(menu);
}

void LauncherApp::DrawWindowBackground(HDC hdc)
{
    if (hdc == nullptr)
    {
        return;
    }

    RECT rect{};
    GetClientRect(mainWindow_, &rect);

    const int width = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    const int contentLeft = 14;
    const int contentWidth = width - 28;
    const int searchCardTop = 14;
    const int searchCardWidth = contentWidth;
    const int searchCardLeft = contentLeft;
    const int searchCardHeight = 42;
    const int resultsCardTop = 64;
    const int resultsCardHeight = 286;

    RECT fullRect{0, 0, width, height};
    const HBRUSH pageBrush = CreateSolidBrush(RGB(246, 244, 240));
    FillRect(hdc, &fullRect, pageBrush);
    DeleteObject(pageBrush);

    auto drawCard = [&](const RECT& sourceRect, COLORREF fill, COLORREF border, int radius)
    {
        RECT shadowRect = sourceRect;
        OffsetRect(&shadowRect, 0, 4);
        const HBRUSH shadowBrush = CreateSolidBrush(RGB(236, 233, 228));
        const HPEN shadowPen = CreatePen(PS_SOLID, 1, RGB(236, 233, 228));
        const HGDIOBJ oldShadowBrush = SelectObject(hdc, shadowBrush);
        const HGDIOBJ oldShadowPen = SelectObject(hdc, shadowPen);
        RoundRect(hdc, shadowRect.left, shadowRect.top, shadowRect.right, shadowRect.bottom, radius, radius);
        SelectObject(hdc, oldShadowPen);
        SelectObject(hdc, oldShadowBrush);
        DeleteObject(shadowPen);
        DeleteObject(shadowBrush);

        const HBRUSH fillBrush = CreateSolidBrush(fill);
        const HPEN borderPen = CreatePen(PS_SOLID, 1, border);
        const HGDIOBJ oldFillBrush = SelectObject(hdc, fillBrush);
        const HGDIOBJ oldBorderPen = SelectObject(hdc, borderPen);
        RoundRect(hdc, sourceRect.left, sourceRect.top, sourceRect.right, sourceRect.bottom, radius, radius);
        SelectObject(hdc, oldBorderPen);
        SelectObject(hdc, oldFillBrush);
        DeleteObject(borderPen);
        DeleteObject(fillBrush);
    };

    RECT searchCard{searchCardLeft, searchCardTop, searchCardLeft + searchCardWidth, searchCardTop + searchCardHeight};
    RECT resultsCard{contentLeft, resultsCardTop, contentLeft + contentWidth, resultsCardTop + resultsCardHeight};

    drawCard(searchCard, RGB(255, 255, 255), RGB(232, 232, 236), 22);
    drawCard(resultsCard, RGB(255, 255, 255), RGB(232, 232, 236), 24);

    RECT footerRect{contentLeft + 14, resultsCard.bottom - 62, resultsCard.right - 14, resultsCard.bottom - 14};
    const HBRUSH footerBrush = CreateSolidBrush(RGB(255, 255, 255));
    FillRect(hdc, &footerRect, footerBrush);
    DeleteObject(footerBrush);
}

void LauncherApp::DrawAutostartToggle(const DRAWITEMSTRUCT* drawItem)
{
    if (drawItem == nullptr)
    {
        return;
    }

    RECT bounds = drawItem->rcItem;
    const HBRUSH boundsBrush = CreateSolidBrush(RGB(255, 255, 255));
    FillRect(drawItem->hDC, &bounds, boundsBrush);
    DeleteObject(boundsBrush);

    RECT track = bounds;
    InflateRect(&track, 0, -2);

    const bool pressed = (drawItem->itemState & ODS_SELECTED) != 0;
    const COLORREF trackColor = autostartEnabled_ ? RGB(52, 168, 83) : RGB(198, 204, 210);
    const COLORREF borderColor = autostartEnabled_ ? RGB(43, 146, 72) : RGB(176, 183, 190);
    const COLORREF knobColor = RGB(255, 255, 255);

    const HBRUSH trackBrush = CreateSolidBrush(trackColor);
    const HPEN trackPen = CreatePen(PS_SOLID, 1, borderColor);
    const HGDIOBJ oldBrush = SelectObject(drawItem->hDC, trackBrush);
    const HGDIOBJ oldPen = SelectObject(drawItem->hDC, trackPen);
    RoundRect(drawItem->hDC, track.left, track.top, track.right, track.bottom, track.bottom - track.top, track.bottom - track.top);
    SelectObject(drawItem->hDC, oldPen);
    SelectObject(drawItem->hDC, oldBrush);
    DeleteObject(trackPen);
    DeleteObject(trackBrush);

    const int inset = 4;
    const int knobSize = (track.bottom - track.top) - inset * 2;
    int knobLeft = autostartEnabled_
        ? track.right - inset - knobSize
        : track.left + inset;
    if (pressed)
    {
        knobLeft += autostartEnabled_ ? -1 : 1;
    }

    const RECT knobRect{knobLeft, track.top + inset, knobLeft + knobSize, track.top + inset + knobSize};
    const HBRUSH knobBrush = CreateSolidBrush(knobColor);
    const HPEN knobPen = CreatePen(PS_SOLID, 1, RGB(214, 218, 224));
    const HGDIOBJ oldKnobBrush = SelectObject(drawItem->hDC, knobBrush);
    const HGDIOBJ oldKnobPen = SelectObject(drawItem->hDC, knobPen);
    Ellipse(drawItem->hDC, knobRect.left, knobRect.top, knobRect.right, knobRect.bottom);
    SelectObject(drawItem->hDC, oldKnobPen);
    SelectObject(drawItem->hDC, oldKnobBrush);
    DeleteObject(knobPen);
    DeleteObject(knobBrush);

    if ((drawItem->itemState & ODS_FOCUS) != 0)
    {
        RECT focusRect = bounds;
        InflateRect(&focusRect, -1, -1);
        DrawFocusRect(drawItem->hDC, &focusRect);
    }
}

void LauncherApp::DrawResultItem(const DRAWITEMSTRUCT* drawItem)
{
    if (drawItem == nullptr || drawItem->itemID == static_cast<UINT>(-1))
    {
        return;
    }

    const bool selected = (drawItem->itemState & ODS_SELECTED) != 0;
    const COLORREF foreground = RGB(28, 28, 28);

    const HBRUSH baseBrush = CreateSolidBrush(RGB(255, 255, 255));
    FillRect(drawItem->hDC, &drawItem->rcItem, baseBrush);
    DeleteObject(baseBrush);

    RECT pillRect = drawItem->rcItem;
    InflateRect(&pillRect, -5, -4);
    if (selected)
    {
        const HBRUSH highlightBrush = CreateSolidBrush(RGB(239, 244, 252));
        const HPEN highlightPen = CreatePen(PS_SOLID, 1, RGB(219, 228, 242));
        const HGDIOBJ oldHighlightBrush = SelectObject(drawItem->hDC, highlightBrush);
        const HGDIOBJ oldHighlightPen = SelectObject(drawItem->hDC, highlightPen);
        RoundRect(drawItem->hDC, pillRect.left, pillRect.top, pillRect.right, pillRect.bottom, 16, 16);
        SelectObject(drawItem->hDC, oldHighlightPen);
        SelectObject(drawItem->hDC, oldHighlightBrush);
        DeleteObject(highlightPen);
        DeleteObject(highlightBrush);
    }
    SetBkMode(drawItem->hDC, TRANSPARENT);
    SetTextColor(drawItem->hDC, foreground);

    LaunchEntry entry;
    int iconIndex = -1;
    {
        ScopedCriticalSection lock(dataLock_);
        if (drawItem->itemID >= currentResults_.size())
        {
            return;
        }

        const auto& result = currentResults_[drawItem->itemID];
        if (result.entryIndex == SIZE_MAX)
        {
            iconIndex = ResolveIconIndex(directEntry_);
            entry = directEntry_;
        }
        else
        {
            auto& actualEntry = (*currentEntries_)[result.entryIndex];
            iconIndex = ResolveIconIndex(actualEntry);
            entry = actualEntry;
        }
    }
    const int iconX = pillRect.left + 10;
    const int iconY = pillRect.top + 8;

    if (shellImageList_ != nullptr && iconIndex >= 0)
    {
        ImageList_Draw(shellImageList_, iconIndex, drawItem->hDC, iconX, iconY, ILD_NORMAL);
    }

    RECT textRect = pillRect;
    textRect.left += 40;
    textRect.right -= 12;
    DrawTextW(drawItem->hDC, BuildDisplayLabel(entry).c_str(), -1, &textRect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);

    if ((drawItem->itemState & ODS_FOCUS) != 0)
    {
        RECT focusRect = pillRect;
        InflateRect(&focusRect, -1, -1);
        DrawFocusRect(drawItem->hDC, &focusRect);
    }
}

void LauncherApp::EnsureShellImageList()
{
    if (shellImageList_ != nullptr)
    {
        return;
    }

    SHFILEINFOW fileInfo{};
    shellImageList_ = reinterpret_cast<HIMAGELIST>(SHGetFileInfoW(
        L"C:\\",
        FILE_ATTRIBUTE_DIRECTORY,
        &fileInfo,
        sizeof(fileInfo),
        SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
}

int LauncherApp::ResolveIconIndex(LaunchEntry& entry)
{
    if (entry.iconIndex >= 0)
    {
        return entry.iconIndex;
    }

    EnsureShellImageList();

    SHFILEINFOW fileInfo{};
    UINT flags = SHGFI_SYSICONINDEX | SHGFI_SMALLICON;

    if (entry.isUrl)
    {
        flags |= SHGFI_USEFILEATTRIBUTES;
        SHGetFileInfoW(L".url", FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(fileInfo), flags);
    }
    else if (entry.isDirectory)
    {
        flags |= SHGFI_USEFILEATTRIBUTES;
        SHGetFileInfoW(entry.fullPath.c_str(), FILE_ATTRIBUTE_DIRECTORY, &fileInfo, sizeof(fileInfo), flags);
    }
    else
    {
        SHGetFileInfoW(entry.fullPath.c_str(), FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(fileInfo), flags);
    }

    entry.iconIndex = fileInfo.iIcon;
    return entry.iconIndex;
}

bool LauncherApp::TryBuildDirectEntry(const std::wstring& query, LaunchEntry& entry) const
{
    const auto trimmed = TrimCopy(query);
    if (trimmed.empty())
    {
        return false;
    }

    if (IsUrlQuery(trimmed))
    {
        auto url = trimmed;
        if (StartsWith(ToLowerCopy(url), L"www."))
        {
            url = L"https://" + url;
        }

        entry = LaunchEntry{};
        entry.name = url;
        entry.fullPath = url;
        entry.sourceTag = L"URL";
        entry.searchNameCompact = NormalizeCompact(entry.name);
        entry.searchPathCompact = NormalizeCompact(entry.fullPath);
        entry.searchInitials = BuildInitials(entry.name);
        entry.isUrl = true;
        return true;
    }

    wchar_t expanded[MAX_PATH * 4] = {};
    const auto stripped = StripWrappingQuotes(trimmed);
    const DWORD expandedLength = ExpandEnvironmentStringsW(stripped.c_str(), expanded, static_cast<DWORD>(std::size(expanded)));
    std::wstring path = expandedLength > 0 && expandedLength < std::size(expanded) ? expanded : stripped;
    path = StripWrappingQuotes(path);

    const DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return false;
    }

    entry = LaunchEntry{};
    entry.fullPath = path;
    entry.isDirectory = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    entry.extension = entry.isDirectory ? L"" : GetExtension(path);
    entry.name = entry.isDirectory ? GetLeafName(path) : GetFileNameWithoutExtension(GetLeafName(path));
    if (entry.name.empty())
    {
        entry.name = GetLeafName(path);
    }
    entry.sourceTag = L"Direct";
    entry.searchNameCompact = NormalizeCompact(entry.name);
    entry.searchPathCompact = NormalizeCompact(path);
    entry.searchInitials = BuildInitials(entry.name);
    return true;
}

bool LauncherApp::IsUrlQuery(const std::wstring& query) const
{
    const auto lower = ToLowerCopy(query);
    return StartsWith(lower, L"http://")
        || StartsWith(lower, L"https://")
        || StartsWith(lower, L"ftp://")
        || StartsWith(lower, L"mailto:")
        || StartsWith(lower, L"www.");
}

bool LauncherApp::TryLaunch(const LaunchEntry& entry, const std::wstring& query)
{
    if (entry.isUrl)
    {
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(mainWindow_, L"open", entry.fullPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL));
        if (result <= 32)
        {
            PublishStatus(L"Launch failed.");
            return false;
        }
    }
    else if (entry.isDirectory)
    {
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(mainWindow_, L"open", entry.fullPath.c_str(), nullptr, entry.fullPath.c_str(), SW_SHOWNORMAL));
        if (result <= 32)
        {
            PublishStatus(L"Launch failed.");
            return false;
        }
    }
    else if (entry.extension == L".ps1")
    {
        std::wstring commandLine = L"powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"" + entry.fullPath + L"\"";
        STARTUPINFOW startupInfo{};
        startupInfo.cb = sizeof(startupInfo);
        PROCESS_INFORMATION processInfo{};
        std::wstring mutableCommand = commandLine;

        const auto workingDirectory = GetDirectoryPart(entry.fullPath);
        if (!CreateProcessW(nullptr, mutableCommand.data(), nullptr, nullptr, FALSE, 0, nullptr, workingDirectory.empty() ? nullptr : workingDirectory.c_str(), &startupInfo, &processInfo))
        {
            PublishStatus(L"Launch failed.");
            return false;
        }

        CloseHandle(processInfo.hThread);
        CloseHandle(processInfo.hProcess);
    }
    else
    {
        const auto directory = GetDirectoryPart(entry.fullPath);
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(mainWindow_, L"open", entry.fullPath.c_str(), nullptr, directory.empty() ? nullptr : directory.c_str(), SW_SHOWNORMAL));
        if (result <= 32)
        {
            PublishStatus(L"Launch failed.");
            return false;
        }
    }

    RecordLaunch(entry.fullPath, query);
    PublishStatus(L"Launched.");
    return true;
}

bool LauncherApp::TryLaunchAsAdmin(const LaunchEntry& entry, const std::wstring& query)
{
    if (!CanRunAsAdmin(entry))
    {
        PublishStatus(L"Run as administrator is unavailable for this item.");
        return false;
    }

    if (entry.extension == L".ps1")
    {
        const std::wstring parameters = L"-NoProfile -ExecutionPolicy Bypass -File \"" + entry.fullPath + L"\"";
        const auto directory = GetDirectoryPart(entry.fullPath);
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(
            mainWindow_,
            L"runas",
            L"powershell.exe",
            parameters.c_str(),
            directory.empty() ? nullptr : directory.c_str(),
            SW_SHOWNORMAL));
        if (result <= 32)
        {
            PublishStatus(L"Administrator launch failed.");
            return false;
        }
    }
    else
    {
        const auto directory = GetDirectoryPart(entry.fullPath);
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(
            mainWindow_,
            L"runas",
            entry.fullPath.c_str(),
            nullptr,
            directory.empty() ? nullptr : directory.c_str(),
            SW_SHOWNORMAL));
        if (result <= 32)
        {
            PublishStatus(L"Administrator launch failed.");
            return false;
        }
    }

    RecordLaunch(entry.fullPath, query);
    PublishStatus(L"Launched as administrator.");
    return true;
}

bool LauncherApp::TryGetSelectedEntry(LaunchEntry& entry) const
{
    const auto selectedIndex = static_cast<int>(SendMessageW(resultsList_, LB_GETCURSEL, 0, 0));
    if (selectedIndex < 0 || static_cast<size_t>(selectedIndex) >= currentResults_.size())
    {
        return false;
    }

    if (currentResults_[selectedIndex].entryIndex == SIZE_MAX)
    {
        entry = directEntry_;
        return true;
    }

    ScopedCriticalSection lock(const_cast<CRITICAL_SECTION&>(dataLock_));
    entry = (*currentEntries_)[currentResults_[selectedIndex].entryIndex];
    return true;
}

bool LauncherApp::OpenEntryLocation(const LaunchEntry& entry)
{
    if (entry.isUrl)
    {
        PublishStatus(L"URL has no filesystem location.");
        return false;
    }

    if (entry.isDirectory)
    {
        const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(mainWindow_, L"open", entry.fullPath.c_str(), nullptr, entry.fullPath.c_str(), SW_SHOWNORMAL));
        if (result <= 32)
        {
            PublishStatus(L"Failed to open location.");
            return false;
        }

        PublishStatus(L"Opened location.");
        return true;
    }

    const std::wstring parameters = L"/select,\"" + entry.fullPath + L"\"";
    const auto result = reinterpret_cast<INT_PTR>(ShellExecuteW(mainWindow_, L"open", L"explorer.exe", parameters.c_str(), nullptr, SW_SHOWNORMAL));
    if (result <= 32)
    {
        PublishStatus(L"Failed to open location.");
        return false;
    }

    PublishStatus(L"Opened location.");
    return true;
}

bool LauncherApp::CopyTextToClipboard(const std::wstring& text)
{
    if (text.empty())
    {
        return false;
    }

    if (!OpenClipboard(mainWindow_))
    {
        return false;
    }

    EmptyClipboard();

    const SIZE_T bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (memory == nullptr)
    {
        CloseClipboard();
        return false;
    }

    void* locked = GlobalLock(memory);
    if (locked == nullptr)
    {
        GlobalFree(memory);
        CloseClipboard();
        return false;
    }

    CopyMemory(locked, text.c_str(), bytes);
    GlobalUnlock(memory);

    if (SetClipboardData(CF_UNICODETEXT, memory) == nullptr)
    {
        GlobalFree(memory);
        CloseClipboard();
        return false;
    }

    CloseClipboard();
    return true;
}

bool LauncherApp::CanRunAsAdmin(const LaunchEntry& entry) const
{
    if (entry.isUrl || entry.isDirectory)
    {
        return false;
    }

    return entry.extension == L".exe"
        || entry.extension == L".bat"
        || entry.extension == L".cmd"
        || entry.extension == L".com"
        || entry.extension == L".ps1"
        || entry.extension == L".vbs"
        || entry.extension == L".vbe"
        || entry.extension == L".js"
        || entry.extension == L".jse"
        || entry.extension == L".wsf"
        || entry.extension == L".msc"
        || entry.extension == L".lnk"
        || entry.extension == L".msi"
        || entry.extension == L".appref-ms";
}

void LauncherApp::StartIndexThread(bool forceRebuild)
{
    bool expected = false;
    if (!indexRunning_.compare_exchange_strong(expected, true))
    {
        PublishStatus(L"Index refresh already running.");
        return;
    }

    if (indexThread_ != nullptr)
    {
        WaitForSingleObject(indexThread_, INFINITE);
        CloseHandle(indexThread_);
        indexThread_ = nullptr;
    }

    auto* args = new IndexThreadArgs{this, forceRebuild};
    indexThread_ = CreateThread(nullptr, 0, IndexThreadProc, args, 0, nullptr);
    if (indexThread_ == nullptr)
    {
        delete args;
        indexRunning_ = false;
        PublishStatus(L"Failed to start index thread.");
    }
}

std::vector<std::wstring> LauncherApp::BuildScanRoots() const
{
    std::vector<std::wstring> roots;

    for (const auto& root : GetFixedDriveRoots())
    {
        AppendUniqueRoot(roots, root);
    }

    ScopedCriticalSection lock(const_cast<CRITICAL_SECTION&>(dataLock_));
    for (const auto& include : settings_.includeDirectories)
    {
        AppendUniqueRoot(roots, include);
    }

    return roots;
}

std::vector<std::wstring> LauncherApp::BuildWatchRoots() const
{
    std::vector<std::wstring> roots;

    for (const auto& root : desktopRoots_)
    {
        AppendUniqueRoot(roots, root);
    }

    for (const auto& root : startMenuRoots_)
    {
        AppendUniqueRoot(roots, root);
    }

    for (const auto& root : pinnedRoots_)
    {
        AppendUniqueRoot(roots, root);
    }

    ScopedCriticalSection lock(const_cast<CRITICAL_SECTION&>(dataLock_));
    for (const auto& include : settings_.includeDirectories)
    {
        AppendUniqueRoot(roots, include);
    }

    for (const auto& priority : settings_.priorityDirectories)
    {
        AppendUniqueRoot(roots, priority);
    }

    return roots;
}

void LauncherApp::IndexWorker(bool forceRebuild)
{
    PublishStatus(forceRebuild ? L"Rebuilding index..." : L"Refreshing index in background...");

    std::vector<LaunchEntry> rebuilt;
    rebuilt.reserve(32768);
    size_t nextStatusThreshold = 5000;
    const auto roots = BuildScanRoots();

    for (const auto& root : roots)
    {
        if (shuttingDown_)
        {
            break;
        }

        PublishStatus((forceRebuild ? L"Scanning " : L"Background scanning ") + root + L" ...");
        ScanRoot(root, rebuilt, nextStatusThreshold);
    }

    std::sort(rebuilt.begin(), rebuilt.end(), [](const LaunchEntry& left, const LaunchEntry& right)
    {
        return _wcsicmp(left.name.c_str(), right.name.c_str()) < 0;
    });

    const auto generatedTicks = CurrentUtcTicks();

    {
        ScopedCriticalSection lock(dataLock_);
        entries_ = std::make_shared<std::vector<LaunchEntry>>(std::move(rebuilt));
        lastIndexedUtcTicks_ = generatedTicks;
    }

    RebuildSearchCaches();
    SaveCache(*entries_, generatedTicks);

    std::wstringstream stream;
    stream << L"Index updated, " << entries_->size() << L" items.";
    PublishStatus(stream.str());
    PostMessageW(mainWindow_, kIndexUpdatedMessage, 0, 0);
    indexRunning_ = false;
}

void LauncherApp::LoadCache()
{
    std::ifstream stream(Utf8FromWide(cachePath_), std::ios::binary);
    if (!stream.is_open())
    {
        PublishStatus(L"No cache found. First launch will build index.");
        return;
    }

    std::vector<LaunchEntry> loadedEntries;
    long long generatedTicks = 0;
    std::string utf8Line;

    while (std::getline(stream, utf8Line))
    {
        if (!utf8Line.empty() && utf8Line.back() == '\r')
        {
            utf8Line.pop_back();
        }

        const auto line = WideFromUtf8(utf8Line);
        const auto fields = SplitEscapedLine(line);
        if (fields.empty())
        {
            continue;
        }

        if (fields[0] == L"V")
        {
            if (fields.size() >= 3)
            {
                generatedTicks = _wcstoi64(fields[2].c_str(), nullptr, 10);
            }

            continue;
        }

        if (fields[0] != L"E" || fields.size() < 8)
        {
            continue;
        }

        LaunchEntry entry;
        entry.fullPath = fields[1];
        entry.name = fields[2];
        entry.extension = fields[3];
        if (fields.size() >= 10)
        {
            entry.sourceTag = fields[4];
            entry.searchNameCompact = fields[5];
            entry.searchPathCompact = fields[6];
            entry.searchInitials = fields[7];
            entry.isDirectory = fields[8] == L"1";
            entry.lastWriteUtcTicks = _wcstoi64(fields[9].c_str(), nullptr, 10);
        }
        else
        {
            entry.sourceTag = BuildSourceTag(entry.fullPath, entry.isDirectory);
            entry.searchNameCompact = fields[4];
            entry.searchPathCompact = fields[5];
            entry.searchInitials = fields[6];
            entry.lastWriteUtcTicks = _wcstoi64(fields[7].c_str(), nullptr, 10);
        }

        entry.sourceTag = ResolveSourceTag(entry);
        loadedEntries.push_back(std::move(entry));
    }

    loadedEntries.erase(
        std::remove_if(loadedEntries.begin(), loadedEntries.end(), [this](const LaunchEntry& entry)
        {
            if (IsExcludedDirectoryLocked(entry.fullPath))
            {
                return true;
            }

            if (entry.isDirectory && !ShouldIndexDirectoryEntryLocked(entry.fullPath))
            {
                return true;
            }

            return false;
        }),
        loadedEntries.end());

    {
        ScopedCriticalSection lock(dataLock_);
        entries_ = std::make_shared<std::vector<LaunchEntry>>(std::move(loadedEntries));
        currentEntries_ = entries_;
        lastIndexedUtcTicks_ = generatedTicks;
    }

    RebuildSearchCaches();

    std::wstringstream streamText;
    streamText << L"Loaded cache, " << (entries_ == nullptr ? 0 : entries_->size()) << L" items.";
    PublishStatus(streamText.str());
}

void LauncherApp::SaveCache(const std::vector<LaunchEntry>& entries, long long generatedUtcTicks)
{
    std::ofstream stream(Utf8FromWide(cachePath_), std::ios::binary | std::ios::trunc);
    if (!stream.is_open())
    {
        return;
    }

    WriteUtf8Line(stream, L"V\t3\t" + std::to_wstring(generatedUtcTicks));

    for (const auto& entry : entries)
    {
        std::wstring line = L"E\t";
        line += EscapeField(entry.fullPath);
        line += L"\t" + EscapeField(entry.name);
        line += L"\t" + EscapeField(entry.extension);
        line += L"\t" + EscapeField(entry.sourceTag);
        line += L"\t" + EscapeField(entry.searchNameCompact);
        line += L"\t" + EscapeField(entry.searchPathCompact);
        line += L"\t" + EscapeField(entry.searchInitials);
        line += L"\t" + std::wstring(entry.isDirectory ? L"1" : L"0");
        line += L"\t" + std::to_wstring(entry.lastWriteUtcTicks);
        WriteUtf8Line(stream, line);
    }
}

void LauncherApp::LoadUsage()
{
    std::ifstream stream(Utf8FromWide(usagePath_), std::ios::binary);
    if (!stream.is_open())
    {
        return;
    }

    std::unordered_map<std::wstring, UsageStat> loadedUsage;
    std::unordered_map<std::wstring, QueryBucket> loadedQueryHistory;
    std::string utf8Line;

    while (std::getline(stream, utf8Line))
    {
        if (!utf8Line.empty() && utf8Line.back() == '\r')
        {
            utf8Line.pop_back();
        }

        const auto line = WideFromUtf8(utf8Line);
        const auto fields = SplitEscapedLine(line);
        if (fields.empty())
        {
            continue;
        }

        if (fields[0] == L"P" && fields.size() >= 4)
        {
            loadedUsage[fields[1]] = UsageStat{_wtoi(fields[2].c_str()), _wcstoi64(fields[3].c_str(), nullptr, 10)};
            continue;
        }

        if (fields[0] == L"Q" && fields.size() >= 5)
        {
            auto& bucket = loadedQueryHistory[fields[1]];
            bucket.lastLaunchUtcTicks = _wcstoi64(fields[2].c_str(), nullptr, 10);
            bucket.launchCounts[fields[3]] = _wtoi(fields[4].c_str());
        }
    }

    ScopedCriticalSection lock(dataLock_);
    usage_ = std::move(loadedUsage);
    queryHistory_ = std::move(loadedQueryHistory);
}

void LauncherApp::SaveUsageSnapshot(
    const std::unordered_map<std::wstring, UsageStat>& usage,
    const std::unordered_map<std::wstring, QueryBucket>& queryHistory)
{
    std::ofstream stream(Utf8FromWide(usagePath_), std::ios::binary | std::ios::trunc);
    if (!stream.is_open())
    {
        return;
    }

    WriteUtf8Line(stream, L"V\t1");

    for (const auto& pair : usage)
    {
        std::wstring line = L"P\t";
        line += EscapeField(pair.first);
        line += L"\t" + std::to_wstring(pair.second.launchCount);
        line += L"\t" + std::to_wstring(pair.second.lastLaunchUtcTicks);
        WriteUtf8Line(stream, line);
    }

    for (const auto& pair : queryHistory)
    {
        for (const auto& pathPair : pair.second.launchCounts)
        {
            std::wstring line = L"Q\t";
            line += EscapeField(pair.first);
            line += L"\t" + std::to_wstring(pair.second.lastLaunchUtcTicks);
            line += L"\t" + EscapeField(pathPair.first);
            line += L"\t" + std::to_wstring(pathPair.second);
            WriteUtf8Line(stream, line);
        }
    }
}

void LauncherApp::RecordLaunch(const std::wstring& fullPath, const std::wstring& query)
{
    const auto now = CurrentUtcTicks();
    const auto queryCompact = NormalizeCompact(query);
    std::shared_ptr<std::vector<LaunchEntry>> entriesSnapshot;
    std::unordered_map<std::wstring, UsageStat> usageSnapshot;
    std::unordered_map<std::wstring, QueryBucket> queryHistorySnapshot;
    std::vector<std::wstring> priorityRoots;

    {
        ScopedCriticalSection lock(dataLock_);
        auto& usage = usage_[fullPath];
        usage.launchCount += 1;
        usage.lastLaunchUtcTicks = now;

        if (!queryCompact.empty())
        {
            auto& bucket = queryHistory_[queryCompact];
            bucket.lastLaunchUtcTicks = now;
            bucket.launchCounts[fullPath] += 1;
            TrimBucket(bucket);

            if (queryHistory_.size() > 400)
            {
                std::vector<std::pair<std::wstring, long long>> keys;
                keys.reserve(queryHistory_.size());
                for (const auto& pair : queryHistory_)
                {
                    keys.emplace_back(pair.first, pair.second.lastLaunchUtcTicks);
                }

                std::sort(keys.begin(), keys.end(), [](const auto& left, const auto& right)
                {
                    return left.second > right.second;
                });

                while (keys.size() > 400)
                {
                    queryHistory_.erase(keys.back().first);
                    keys.pop_back();
                }
            }
        }

        entriesSnapshot = entries_;
        usageSnapshot = usage_;
        queryHistorySnapshot = queryHistory_;
        priorityRoots = settings_.priorityDirectories;
    }

    SaveUsageSnapshot(usageSnapshot, queryHistorySnapshot);

    auto emptyQueryCandidates = BuildEmptyQueryCandidateIndices(entriesSnapshot, usageSnapshot, priorityRoots);
    {
        ScopedCriticalSection lock(dataLock_);
        if (entries_ == entriesSnapshot)
        {
            emptyQueryCandidateIndices_ = std::move(emptyQueryCandidates);
        }
    }
}

std::unordered_map<std::wstring, int> LauncherApp::BuildQueryAffinityLocked(const std::wstring& queryCompact) const
{
    return BuildQueryAffinitySnapshot(queryHistory_, queryCompact);
}

std::unordered_map<std::wstring, int> LauncherApp::BuildQueryAffinitySnapshot(
    const std::unordered_map<std::wstring, QueryBucket>& queryHistory,
    const std::wstring& queryCompact)
{
    std::unordered_map<std::wstring, int> result;
    if (queryCompact.empty())
    {
        return result;
    }

    for (const auto& pair : queryHistory)
    {
        const int weight = GetQueryRelationWeight(queryCompact, pair.first);
        if (weight == 0)
        {
            continue;
        }

        for (const auto& pathPair : pair.second.launchCounts)
        {
            result[pathPair.first] += pathPair.second * weight;
        }
    }

    return result;
}

int LauncherApp::GetQueryRelationWeight(const std::wstring& currentQuery, const std::wstring& storedQuery)
{
    if (currentQuery == storedQuery)
    {
        return 7;
    }

    if (StartsWith(storedQuery, currentQuery) || StartsWith(currentQuery, storedQuery))
    {
        return 3;
    }

    if (storedQuery.find(currentQuery) != std::wstring::npos || currentQuery.find(storedQuery) != std::wstring::npos)
    {
        return 1;
    }

    return 0;
}

int LauncherApp::CalculateScore(
    const LaunchEntry& entry,
    const std::wstring& queryCompact,
    const std::vector<std::wstring>& queryTokens,
    const UsageStat* usage,
    int queryAffinity,
    int priorityBoost)
{
    const int extensionBoost = ExtensionBoost(entry.extension);
    const int categoryBoost = GetEntryCategoryBoost(entry, queryCompact.size());
    const int usageBoost = std::min((usage == nullptr ? 0 : usage->launchCount) * 90, 1200);
    const int recencyBoost = CalculateRecencyBoost(usage == nullptr ? 0 : usage->lastLaunchUtcTicks);

    if (queryCompact.empty())
    {
        return 900 + usageBoost + recencyBoost + extensionBoost + categoryBoost + priorityBoost
            - std::min(static_cast<int>(entry.name.size()) * 3, 180);
    }

    const int nameScore = ScorePattern(queryCompact, entry.searchNameCompact);
    const int initialsScore = ScorePattern(queryCompact, entry.searchInitials);
    const int pathScore = ScorePattern(queryCompact, entry.searchPathCompact) / 4;
    const bool hasTextMatch = nameScore > 0 || initialsScore > 0 || pathScore > 0;
    int tokenBonus = 0;

    for (const auto& token : queryTokens)
    {
        const int tokenScore = std::max(
            ScorePattern(token, entry.searchNameCompact),
            std::max(ScorePattern(token, entry.searchInitials), ScorePattern(token, entry.searchPathCompact) / 4));

        if (tokenScore == 0 && queryAffinity == 0)
        {
            return 0;
        }

        tokenBonus += tokenScore / 5;
    }

    if (!hasTextMatch && queryAffinity == 0)
    {
        return 0;
    }

    const int affinityBoost = std::min(queryAffinity * 320, 2600);
    int score = nameScore * 2 + initialsScore * 2 + pathScore + tokenBonus + affinityBoost + usageBoost + recencyBoost
        + extensionBoost + categoryBoost + priorityBoost;
    score -= std::min(static_cast<int>(entry.name.size()) * 4, 220);
    return score;
}

std::vector<size_t> LauncherApp::BuildCandidateIndicesLocked(
    const std::wstring& query,
    const std::wstring& queryCompact,
    const std::vector<std::wstring>& queryTokens,
    const std::shared_ptr<std::vector<LaunchEntry>>& entriesSnapshot) const
{
    constexpr size_t kReusableCandidateLimit = 60000;

    if (entriesSnapshot == nullptr || entriesSnapshot->empty())
    {
        return {};
    }

    if (entriesSnapshot != searchCacheEntries_)
    {
        return {};
    }

    if (queryCompact.empty())
    {
        return emptyQueryCandidateIndices_;
    }

    if (query.find(L'\\') != std::wstring::npos
        || query.find(L'/') != std::wstring::npos
        || query.find(L':') != std::wstring::npos)
    {
        return {};
    }

    if (entriesSnapshot == readySearchEntries_
        && !readySearchQueryCompact_.empty()
        && queryCompact.size() > readySearchQueryCompact_.size()
        && StartsWith(queryCompact, readySearchQueryCompact_)
        && !readySearchMatchedIndices_.empty()
        && readySearchMatchedIndices_.size() <= kReusableCandidateLimit)
    {
        return readySearchMatchedIndices_;
    }

    std::unordered_set<std::wstring> keys;
    AppendPrefixKeys(keys, queryCompact);
    for (const auto& token : queryTokens)
    {
        AppendPrefixKeys(keys, token);
    }

    if (keys.empty())
    {
        return {};
    }

    std::vector<size_t> candidates;
    std::unordered_set<size_t> seen;

    for (const auto& key : keys)
    {
        if (key.size() >= 2)
        {
            AppendIndexedCandidates(candidates, seen, searchPrefixIndex_, key);
        }
    }

    if (candidates.empty())
    {
        for (const auto& key : keys)
        {
            if (key.size() == 1)
            {
                AppendIndexedCandidates(candidates, seen, searchPrefixIndex_, key);
            }
        }
    }

    return candidates;
}

std::vector<size_t> LauncherApp::BuildEmptyQueryCandidateIndices(
    const std::shared_ptr<std::vector<LaunchEntry>>& entriesSnapshot,
    const std::unordered_map<std::wstring, UsageStat>& usageSnapshot,
    const std::vector<std::wstring>& priorityRoots) const
{
    constexpr size_t kHomeCandidateLimit = 192;

    std::vector<size_t> candidates;
    if (entriesSnapshot == nullptr || entriesSnapshot->empty())
    {
        return candidates;
    }

    std::vector<SearchResult> ranked;
    const auto& entries = *entriesSnapshot;
    ranked.reserve(kHomeCandidateLimit);

    for (size_t index = 0; index < entries.size(); ++index)
    {
        const auto& entry = entries[index];
        const auto usageIt = usageSnapshot.find(entry.fullPath);
        const UsageStat* usage = usageIt == usageSnapshot.end() ? nullptr : &usageIt->second;
        const auto normalizedPath = NormalizeDirectoryPath(entry.fullPath);
        int priorityBoost = GetLaunchSurfaceBoost(entry);
        for (const auto& priorityRoot : priorityRoots)
        {
            if (IsSameOrChildPath(normalizedPath, priorityRoot))
            {
                priorityBoost += 900;
                break;
            }
        }

        const int score = CalculateScore(entry, L"", {}, usage, 0, priorityBoost);
        if (score <= 0)
        {
            continue;
        }

        InsertRanked(ranked, SearchResult{index, score, usage == nullptr ? 0 : usage->launchCount}, kHomeCandidateLimit, entries);
    }

    candidates.reserve(ranked.size());
    for (const auto& result : ranked)
    {
        candidates.push_back(result.entryIndex);
    }

    return candidates;
}

std::vector<SearchResult> LauncherApp::BuildResultsForQuery(
    const std::wstring& query,
    const std::shared_ptr<std::vector<LaunchEntry>>& entriesSnapshot,
    const std::unordered_map<std::wstring, UsageStat>& usageSnapshot,
    const std::unordered_map<std::wstring, QueryBucket>& queryHistorySnapshot,
    const std::vector<std::wstring>& excludedRoots,
    const std::vector<std::wstring>& priorityRoots,
    unsigned long long requestId,
    const std::vector<size_t>* candidateIndices,
    std::vector<size_t>* matchedIndices) const
{
    constexpr size_t kVisibleResultLimit = 32;
    constexpr size_t kCandidateResultLimit = 64;
    constexpr size_t kReusableMatchLimit = 40000;

    std::vector<SearchResult> ranked;
    if (entriesSnapshot == nullptr || entriesSnapshot->empty())
    {
        return ranked;
    }

    const auto queryCompact = NormalizeCompact(query);
    const auto queryTokens = TokenizeQuery(query);
    const auto queryAffinity = BuildQueryAffinitySnapshot(queryHistorySnapshot, queryCompact);
    const auto& entries = *entriesSnapshot;
    size_t processedCount = 0;
    bool reusableOverflow = false;

    if (matchedIndices != nullptr)
    {
        matchedIndices->clear();
    }

    const auto processEntry = [&](size_t index) -> bool
    {
        if (index >= entries.size())
        {
            return true;
        }

        if (requestId != 0 && (processedCount % 2048) == 0 && latestSearchRequestId_.load() != requestId)
        {
            return false;
        }

        ++processedCount;

        const auto& entry = entries[index];
        const auto normalizedPath = NormalizeDirectoryPath(entry.fullPath);
        bool excluded = false;
        for (const auto& excludedRoot : excludedRoots)
        {
            if (IsSameOrChildPath(normalizedPath, excludedRoot))
            {
                excluded = true;
                break;
            }
        }

        if (excluded)
        {
            return true;
        }

        const auto usageIt = usageSnapshot.find(entry.fullPath);
        const UsageStat* usage = usageIt == usageSnapshot.end() ? nullptr : &usageIt->second;
        const auto affinityIt = queryAffinity.find(entry.fullPath);
        const int affinity = affinityIt == queryAffinity.end() ? 0 : affinityIt->second;
        int priorityBoost = GetLaunchSurfaceBoost(entry);
        for (const auto& priorityRoot : priorityRoots)
        {
            if (IsSameOrChildPath(normalizedPath, priorityRoot))
            {
                priorityBoost += 900;
                break;
            }
        }

        const int score = CalculateScore(entry, queryCompact, queryTokens, usage, affinity, priorityBoost);
        if (score <= 0)
        {
            return true;
        }

        if (matchedIndices != nullptr && !reusableOverflow)
        {
            if (matchedIndices->size() < kReusableMatchLimit)
            {
                matchedIndices->push_back(index);
            }
            else
            {
                matchedIndices->clear();
                reusableOverflow = true;
            }
        }

        InsertRanked(ranked, SearchResult{index, score, usage == nullptr ? 0 : usage->launchCount}, kCandidateResultLimit, entries);
        return true;
    };

    if (candidateIndices != nullptr && !candidateIndices->empty())
    {
        for (const auto index : *candidateIndices)
        {
            if (!processEntry(index))
            {
                return {};
            }
        }
    }
    else
    {
        for (size_t index = 0; index < entries.size(); ++index)
        {
            if (!processEntry(index))
            {
                return {};
            }
        }
    }

    std::vector<SearchResult> filtered;
    filtered.reserve(std::min(kVisibleResultLimit, ranked.size()));
    std::unordered_set<std::wstring> seenPreferredNames;

    for (const auto& result : ranked)
    {
        const auto& entry = entries[result.entryIndex];
        if (!entry.searchNameCompact.empty())
        {
            if (IsPreferredLaunchSurface(entry))
            {
                if (!seenPreferredNames.insert(entry.searchNameCompact).second)
                {
                    continue;
                }
            }
            else if (seenPreferredNames.find(entry.searchNameCompact) != seenPreferredNames.end())
            {
                continue;
            }
        }

        filtered.push_back(result);
        if (filtered.size() >= kVisibleResultLimit)
        {
            break;
        }
    }

    return filtered;
}

void LauncherApp::InsertRanked(std::vector<SearchResult>& results, const SearchResult& candidate, size_t limit, const std::vector<LaunchEntry>& entries)
{
    if (results.size() == limit && CompareResults(candidate, results.back(), entries) >= 0)
    {
        return;
    }

    auto insertAt = results.end();
    for (auto it = results.begin(); it != results.end(); ++it)
    {
        if (CompareResults(candidate, *it, entries) < 0)
        {
            insertAt = it;
            break;
        }
    }

    results.insert(insertAt, candidate);
    if (results.size() > limit)
    {
        results.pop_back();
    }
}

int LauncherApp::CompareResults(const SearchResult& left, const SearchResult& right, const std::vector<LaunchEntry>& entries)
{
    if (left.score != right.score)
    {
        return right.score - left.score;
    }

    if (left.usageCount != right.usageCount)
    {
        return right.usageCount - left.usageCount;
    }

    const auto& leftEntry = entries[left.entryIndex];
    const auto& rightEntry = entries[right.entryIndex];
    const int leftRank = GetEntryResultRank(leftEntry);
    const int rightRank = GetEntryResultRank(rightEntry);

    if (leftRank != rightRank)
    {
        return leftRank - rightRank;
    }

    if (leftEntry.name.size() != rightEntry.name.size())
    {
        return static_cast<int>(leftEntry.name.size()) - static_cast<int>(rightEntry.name.size());
    }

    return _wcsicmp(leftEntry.name.c_str(), rightEntry.name.c_str());
}

void LauncherApp::ScanRoot(const std::wstring& root, std::vector<LaunchEntry>& results, size_t& nextStatusThreshold)
{
    std::vector<std::wstring> pending;
    pending.push_back(root);

    while (!pending.empty())
    {
        const auto currentDirectory = pending.back();
        pending.pop_back();

        if (ShouldSkipDirectory(currentDirectory) || IsExcludedDirectoryLocked(currentDirectory))
        {
            continue;
        }

        WIN32_FIND_DATAW data{};
        const auto searchPath = JoinPath(currentDirectory, L"*");
        const HANDLE findHandle = FindFirstFileW(searchPath.c_str(), &data);
        if (findHandle == INVALID_HANDLE_VALUE)
        {
            continue;
        }

        do
        {
            const std::wstring fileName = data.cFileName;
            if (fileName == L"." || fileName == L"..")
            {
                continue;
            }

            const auto fullPath = JoinPath(currentDirectory, fileName);

            if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
                if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0
                    && !ShouldSkipDirectory(fullPath)
                    && !IsExcludedDirectoryLocked(fullPath))
                {
                    if (ShouldIndexDirectoryEntryLocked(fullPath))
                    {
                        LaunchEntry directoryEntry;
                        if (TryCreateEntry(fullPath, data, directoryEntry))
                        {
                            directoryEntry.sourceTag = ResolveSourceTag(directoryEntry);
                            results.push_back(std::move(directoryEntry));
                            if (results.size() >= nextStatusThreshold)
                            {
                                std::wstringstream stream;
                                stream << L"Indexing... found " << results.size() << L" items.";
                                PublishStatus(stream.str());
                                nextStatusThreshold += 5000;
                            }
                        }
                    }

                    pending.push_back(fullPath);
                }

                continue;
            }

            if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
            {
                continue;
            }

            LaunchEntry entry;
            if (!TryCreateEntry(fullPath, data, entry))
            {
                continue;
            }

            entry.sourceTag = ResolveSourceTag(entry);
            results.push_back(std::move(entry));
            if (results.size() >= nextStatusThreshold)
            {
                std::wstringstream stream;
                stream << L"Indexing... found " << results.size() << L" items.";
                PublishStatus(stream.str());
                nextStatusThreshold += 5000;
            }
        }
        while (FindNextFileW(findHandle, &data));

        FindClose(findHandle);
    }
}

bool LauncherApp::TryCreateEntry(const std::wstring& fullPath, const WIN32_FIND_DATAW& data, LaunchEntry& entry)
{
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        entry.fullPath = fullPath;
        entry.name = data.cFileName;
        entry.extension = L"";
        entry.sourceTag = BuildSourceTag(fullPath, true);
        entry.searchNameCompact = NormalizeCompact(entry.name);
        entry.searchPathCompact = NormalizeCompact(fullPath);
        entry.searchInitials = BuildInitials(entry.name);
        entry.isDirectory = true;
        entry.lastWriteUtcTicks = FileTimeToTicks(data.ftLastWriteTime);
        return !entry.name.empty();
    }

    const auto extension = GetExtension(data.cFileName);
    if (!IsLaunchableExtension(extension))
    {
        return false;
    }

    const auto name = GetFileNameWithoutExtension(data.cFileName);
    entry.fullPath = fullPath;
    entry.name = name.empty() ? data.cFileName : name;
    entry.extension = extension;
    entry.sourceTag = BuildSourceTag(fullPath, false);
    entry.searchNameCompact = NormalizeCompact(entry.name);
    entry.searchPathCompact = NormalizeCompact(fullPath);
    entry.searchInitials = BuildInitials(entry.name);
    entry.isDirectory = false;
    entry.lastWriteUtcTicks = FileTimeToTicks(data.ftLastWriteTime);
    return true;
}

bool LauncherApp::IsExcludedDirectoryLocked(const std::wstring& path) const
{
    const auto normalized = NormalizeDirectoryPath(path);
    ScopedCriticalSection lock(const_cast<CRITICAL_SECTION&>(dataLock_));
    for (const auto& excluded : settings_.excludeDirectories)
    {
        if (IsSameOrChildPath(normalized, excluded))
        {
            return true;
        }
    }

    return false;
}

bool LauncherApp::ShouldIndexDirectoryEntryLocked(const std::wstring& path) const
{
    const auto normalized = NormalizeDirectoryPath(path);
    ScopedCriticalSection lock(const_cast<CRITICAL_SECTION&>(dataLock_));

    if (!userProfilePath_.empty() && IsSameOrChildPath(normalized, userProfilePath_))
    {
        return true;
    }

    for (const auto& include : settings_.includeDirectories)
    {
        if (IsSameOrChildPath(normalized, include))
        {
            return true;
        }
    }

    for (const auto& priority : settings_.priorityDirectories)
    {
        if (IsSameOrChildPath(normalized, priority))
        {
            return true;
        }
    }

    return false;
}

int LauncherApp::GetPriorityBoostLocked(const std::wstring& path) const
{
    const auto normalized = NormalizeDirectoryPath(path);
    ScopedCriticalSection lock(const_cast<CRITICAL_SECTION&>(dataLock_));
    for (const auto& priority : settings_.priorityDirectories)
    {
        if (IsSameOrChildPath(normalized, priority))
        {
            return 900;
        }
    }

    return 0;
}

bool LauncherApp::IsUnderAnyRoot(const std::wstring& path, const std::vector<std::wstring>& roots) const
{
    for (const auto& root : roots)
    {
        if (IsSameOrChildPath(path, root))
        {
            return true;
        }
    }

    return false;
}

int LauncherApp::GetLaunchSurfaceBoost(const LaunchEntry& entry) const
{
    if (entry.isUrl)
    {
        return 0;
    }

    const auto normalized = NormalizeDirectoryPath(entry.fullPath);
    const int folderBoost = entry.isDirectory ? 40 : 0;
    const int linkBoost = entry.extension == L".lnk" ? 120 : 0;

    if (IsUnderAnyRoot(normalized, pinnedRoots_))
    {
        return folderBoost + (entry.extension == L".lnk" ? 820 : 700);
    }

    if (IsUnderAnyRoot(normalized, desktopRoots_))
    {
        return folderBoost + (entry.extension == L".lnk" ? 700 : 580);
    }

    if (IsUnderAnyRoot(normalized, startMenuRoots_))
    {
        return folderBoost + (entry.extension == L".lnk" ? 640 : 520);
    }

    return linkBoost;
}

std::wstring LauncherApp::ResolveSourceTag(const LaunchEntry& entry) const
{
    if (entry.isUrl)
    {
        return L"URL";
    }

    const auto normalized = NormalizeDirectoryPath(entry.fullPath);
    if (IsUnderAnyRoot(normalized, pinnedRoots_))
    {
        return L"Pinned";
    }

    if (IsUnderAnyRoot(normalized, desktopRoots_))
    {
        return L"Desktop";
    }

    if (IsUnderAnyRoot(normalized, startMenuRoots_))
    {
        return L"Start Menu";
    }

    return BuildSourceTag(entry.fullPath, entry.isDirectory);
}

void LauncherApp::PublishStatus(const std::wstring& status)
{
    {
        ScopedCriticalSection lock(dataLock_);
        statusText_ = status;
    }

    if (mainWindow_ != nullptr)
    {
        PostMessageW(mainWindow_, kStatusUpdatedMessage, 0, 0);
    }
}

std::wstring LauncherApp::ReadControlText(HWND window) const
{
    if (window == nullptr)
    {
        return L"";
    }

    const int length = GetWindowTextLengthW(window);
    if (length == 0)
    {
        return L"";
    }

    std::wstring text(length + 1, L'\0');
    GetWindowTextW(window, text.data(), length + 1);
    text.resize(length);
    return text;
}

std::wstring LauncherApp::GetLocalAppDataPath() const
{
    wchar_t buffer[MAX_PATH] = {};
    SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, SHGFP_TYPE_CURRENT, buffer);
    return buffer;
}

std::wstring LauncherApp::GetExecutablePath() const
{
    wchar_t buffer[MAX_PATH] = {};
    GetModuleFileNameW(nullptr, buffer, MAX_PATH);
    return buffer;
}
}  // namespace quicklauncher

int APIENTRY wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int)
{
    quicklauncher::LauncherApp app;
    if (!app.Initialize(instance))
    {
        MessageBoxW(nullptr, L"Failed to initialize QuickLauncher Native.", L"QuickLauncher Native", MB_ICONERROR | MB_OK);
        return 1;
    }

    return app.Run();
}
