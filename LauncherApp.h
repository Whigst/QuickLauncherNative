#pragma once

#include <windows.h>
#include <commctrl.h>
#include <shellapi.h>

#include <atomic>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

#include "SearchUtils.h"

namespace quicklauncher
{
constexpr UINT kTrayMessage = WM_APP + 10;
constexpr UINT kToggleLauncherMessage = WM_APP + 11;
constexpr UINT kIndexUpdatedMessage = WM_APP + 12;
constexpr UINT kStatusUpdatedMessage = WM_APP + 13;
constexpr UINT kFilesystemChangedMessage = WM_APP + 14;
constexpr UINT kSearchResultsReadyMessage = WM_APP + 15;
constexpr int kLauncherHotkeyId = 1;

constexpr UINT_PTR kFocusTimerId = 1;
constexpr UINT_PTR kFilesystemDebounceTimerId = 2;
constexpr UINT_PTR kQueryDebounceTimerId = 3;

constexpr int kCommandTrayOpen = 1001;
constexpr int kCommandTrayRebuild = 1002;
constexpr int kCommandTrayExit = 1003;
constexpr int kCommandTrayOpenSettings = 1004;
constexpr int kCommandTrayReloadSettings = 1005;
constexpr int kCommandTrayToggleAutostart = 1006;
constexpr int kCommandResultLaunch = 1101;
constexpr int kCommandResultRunAsAdmin = 1102;
constexpr int kCommandResultOpenLocation = 1103;
constexpr int kCommandResultCopyPath = 1104;

constexpr int kControlQuery = 2001;
constexpr int kControlResults = 2002;
constexpr int kControlPath = 2003;
constexpr int kControlStatus = 2004;
constexpr int kControlAutostartToggle = 2005;

struct LauncherSettings
{
    std::vector<std::wstring> includeDirectories;
    std::vector<std::wstring> excludeDirectories;
    std::vector<std::wstring> priorityDirectories;
    UINT hotkeyModifiers = MOD_ALT;
    UINT hotkeyVirtualKey = VK_SPACE;
    std::wstring hotkeyLabel = L"Alt+Space";
};

class LauncherApp
{
public:
    LauncherApp() = default;
    ~LauncherApp();

    bool Initialize(HINSTANCE instance);
    int Run();

    LRESULT HandleMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
    LRESULT HandleEditMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
    LRESULT HandleListMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam);

private:
    static LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
    static DWORD WINAPI IndexThreadProc(LPVOID param);
    static DWORD WINAPI WatchThreadProc(LPVOID param);
    static DWORD WINAPI SearchThreadProc(LPVOID param);
    static LRESULT CALLBACK EditSubclassProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR id, DWORD_PTR refData);
    static LRESULT CALLBACK ListSubclassProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR id, DWORD_PTR refData);

    void CreateFonts();
    void CreateControls();
    void LayoutControls();
    void CreateTrayIcon();
    void ShowTrayMenu();
    void EnsureSettingsFileExists() const;
    void LoadSettings();
    void OpenSettingsFile() const;
    void ReloadSettingsAndRefresh();
    bool IsAutostartEnabled() const;
    bool SetAutostartEnabled(bool enabled) const;
    void RegisterLauncherHotkey();
    void UnregisterLauncherHotkey();
    void StartWatchThread();
    void StopWatchThread();
    void RestartWatchThread();
    void WatchLoop();
    void StartSearchThread();
    void StopSearchThread();
    void SearchLoop();
    void InitializePreferredRoots();
    void UpdateHintLabel();

    void ShowLauncher();
    void HideLauncher();
    void ToggleLauncher();
    void ReinforceFocus();
    void FocusQueryBox(bool selectAll);
    void PositionWindow();

    void UpdateResults();
    void ScheduleResultsUpdate();
    void RequestSearchUpdate();
    void FlushPendingResultsUpdate();
    void ApplyReadySearchResults();
    void RebuildSearchCaches();
    void PopulateResultsList();
    void UpdateSelectedPath();
    void UpdateFooter();
    void RefreshAutostartToggle();
    void ToggleAutostartSetting();
    void MoveSelection(int offset);
    void LaunchSelected();
    void LaunchSelectedAsAdmin();
    void OpenSelectedLocation();
    void CopySelectedPath();
    void ShowSelectedContextMenu(POINT screenPoint, bool fromKeyboard);
    void DrawWindowBackground(HDC hdc);
    void DrawResultItem(const DRAWITEMSTRUCT* drawItem);
    void DrawAutostartToggle(const DRAWITEMSTRUCT* drawItem);
    void EnsureShellImageList();
    int ResolveIconIndex(LaunchEntry& entry);
    bool TryBuildDirectEntry(const std::wstring& query, LaunchEntry& entry) const;
    bool IsUrlQuery(const std::wstring& query) const;
    bool TryLaunch(const LaunchEntry& entry, const std::wstring& query);
    bool TryLaunchAsAdmin(const LaunchEntry& entry, const std::wstring& query);
    bool TryGetSelectedEntry(LaunchEntry& entry) const;
    bool OpenEntryLocation(const LaunchEntry& entry);
    bool CopyTextToClipboard(const std::wstring& text);
    bool CanRunAsAdmin(const LaunchEntry& entry) const;

    void StartIndexThread(bool forceRebuild);
    void IndexWorker(bool forceRebuild);
    void LoadCache();
    void SaveCache(const std::vector<LaunchEntry>& entries, long long generatedUtcTicks);
    void LoadUsage();
    void SaveUsageSnapshot(
        const std::unordered_map<std::wstring, UsageStat>& usage,
        const std::unordered_map<std::wstring, QueryBucket>& queryHistory);
    void RecordLaunch(const std::wstring& fullPath, const std::wstring& query);
    std::unordered_map<std::wstring, int> BuildQueryAffinityLocked(const std::wstring& queryCompact) const;
    static std::unordered_map<std::wstring, int> BuildQueryAffinitySnapshot(
        const std::unordered_map<std::wstring, QueryBucket>& queryHistory,
        const std::wstring& queryCompact);
    static int GetQueryRelationWeight(const std::wstring& currentQuery, const std::wstring& storedQuery);
    static int CalculateScore(
        const LaunchEntry& entry,
        const std::wstring& queryCompact,
        const std::vector<std::wstring>& queryTokens,
        const UsageStat* usage,
        int queryAffinity,
        int priorityBoost,
        const std::wstring& preferredExtension,
        const std::wstring& preferredDirectory);
    std::vector<size_t> BuildCandidateIndicesLocked(
        const std::wstring& query,
        const std::wstring& queryCompact,
        const std::vector<std::wstring>& queryTokens,
        const std::shared_ptr<std::vector<LaunchEntry>>& entriesSnapshot) const;
    std::vector<size_t> BuildEmptyQueryCandidateIndices(
        const std::shared_ptr<std::vector<LaunchEntry>>& entriesSnapshot,
        const std::unordered_map<std::wstring, UsageStat>& usageSnapshot,
        const std::vector<std::wstring>& priorityRoots) const;
    std::vector<SearchResult> BuildResultsForQuery(
        const std::wstring& query,
        const std::shared_ptr<std::vector<LaunchEntry>>& entriesSnapshot,
        const std::unordered_map<std::wstring, UsageStat>& usageSnapshot,
        const std::unordered_map<std::wstring, QueryBucket>& queryHistorySnapshot,
        const std::vector<std::wstring>& excludedRoots,
        const std::vector<std::wstring>& priorityRoots,
        const std::wstring& preferredExtension,
        const std::wstring& preferredDirectory,
        unsigned long long requestId,
        const std::vector<size_t>* candidateIndices,
        std::vector<size_t>* matchedIndices) const;
    static void InsertRanked(std::vector<SearchResult>& results, const SearchResult& candidate, size_t limit, const std::vector<LaunchEntry>& entries);
    static int CompareResults(const SearchResult& left, const SearchResult& right, const std::vector<LaunchEntry>& entries);
    std::shared_ptr<std::vector<LaunchEntry>> BuildDirectoryDirectiveEntries(const std::wstring& preferredDirectory);

    void ScanRoot(const std::wstring& root, std::vector<LaunchEntry>& results, size_t& nextStatusThreshold);
    static bool TryCreateEntry(const std::wstring& fullPath, const WIN32_FIND_DATAW& data, LaunchEntry& entry);
    bool IsExcludedDirectoryLocked(const std::wstring& path) const;
    bool ShouldIndexDirectoryEntryLocked(const std::wstring& path) const;
    int GetPriorityBoostLocked(const std::wstring& path) const;
    int GetLaunchSurfaceBoost(const LaunchEntry& entry) const;
    std::wstring ResolveSourceTag(const LaunchEntry& entry) const;
    void PublishStatus(const std::wstring& status);
    std::vector<std::wstring> BuildScanRoots() const;
    std::vector<std::wstring> BuildWatchRoots() const;
    bool IsUnderAnyRoot(const std::wstring& path, const std::vector<std::wstring>& roots) const;

    std::wstring ReadControlText(HWND window) const;
    std::wstring GetLocalAppDataPath() const;
    std::wstring GetExecutablePath() const;

    HINSTANCE instance_ = nullptr;
    HWND mainWindow_ = nullptr;
    HWND titleLabel_ = nullptr;
    HWND hintLabel_ = nullptr;
    HWND queryEdit_ = nullptr;
    HWND resultsList_ = nullptr;
    HWND pathLabel_ = nullptr;
    HWND statusLabel_ = nullptr;
    HWND mainTitleLabel_ = nullptr;
    HWND mainSubtitleLabel_ = nullptr;
    HWND sidebarChipLabel_ = nullptr;
    HWND sidebarFooterLabel_ = nullptr;
    HWND autostartLabel_ = nullptr;
    HWND autostartToggle_ = nullptr;

    HFONT titleFont_ = nullptr;
    HFONT normalFont_ = nullptr;
    HFONT smallFont_ = nullptr;
    HIMAGELIST shellImageList_ = nullptr;

    NOTIFYICONDATAW trayIconData_{};

    std::atomic<bool> shuttingDown_ = false;
    std::atomic<bool> indexRunning_ = false;
    std::atomic<bool> watchRefreshQueued_ = false;
    std::atomic<unsigned long long> latestSearchRequestId_ = 0;
    bool queryUpdatePending_ = false;
    int focusRetryCount_ = 0;

    CRITICAL_SECTION dataLock_{};
    std::shared_ptr<std::vector<LaunchEntry>> entries_ = std::make_shared<std::vector<LaunchEntry>>();
    std::shared_ptr<std::vector<LaunchEntry>> currentEntries_ = std::make_shared<std::vector<LaunchEntry>>();
    std::unordered_map<std::wstring, UsageStat> usage_;
    std::unordered_map<std::wstring, QueryBucket> queryHistory_;
    std::vector<SearchResult> currentResults_;
    std::shared_ptr<std::vector<LaunchEntry>> searchCacheEntries_;
    std::unordered_map<std::wstring, std::vector<size_t>> searchPrefixIndex_;
    std::vector<size_t> emptyQueryCandidateIndices_;
    bool searchRequestPending_ = false;
    unsigned long long pendingSearchRequestId_ = 0;
    std::wstring pendingSearchQuery_;
    std::wstring pendingSearchPreferredExtension_;
    std::wstring pendingSearchPreferredDirectory_;
    std::shared_ptr<std::vector<LaunchEntry>> pendingSearchEntries_;
    std::unordered_map<std::wstring, UsageStat> pendingSearchUsage_;
    std::unordered_map<std::wstring, QueryBucket> pendingSearchHistory_;
    std::vector<std::wstring> pendingSearchExcludedRoots_;
    std::vector<std::wstring> pendingSearchPriorityRoots_;
    bool pendingSearchHasDirectEntry_ = false;
    LaunchEntry pendingSearchDirectEntry_;
    std::vector<size_t> pendingSearchCandidateIndices_;
    unsigned long long readySearchRequestId_ = 0;
    std::shared_ptr<std::vector<LaunchEntry>> readySearchEntries_;
    std::vector<SearchResult> readySearchResults_;
    bool readySearchHasDirectEntry_ = false;
    LaunchEntry readySearchDirectEntry_;
    std::wstring readySearchQueryCompact_;
    std::vector<size_t> readySearchMatchedIndices_;
    std::wstring directoryBrowseCacheRoot_;
    std::vector<LaunchEntry> directoryBrowseCacheEntries_;
    bool hasDirectEntry_ = false;
    LaunchEntry directEntry_;
    std::wstring statusText_ = L"Preparing cache...";
    long long lastIndexedUtcTicks_ = 0;

    std::wstring dataDirectory_;
    std::wstring cachePath_;
    std::wstring usagePath_;
    std::wstring settingsPath_;
    std::wstring userProfilePath_;
    std::wstring activeHotkeyLabel_ = L"Alt+Space";
    bool autostartEnabled_ = false;
    std::vector<std::wstring> desktopRoots_;
    std::vector<std::wstring> startMenuRoots_;
    std::vector<std::wstring> pinnedRoots_;
    LauncherSettings settings_;

    HANDLE indexThread_ = nullptr;
    HANDLE watchThread_ = nullptr;
    HANDLE watchStopEvent_ = nullptr;
    HANDLE searchThread_ = nullptr;
    HANDLE searchWakeEvent_ = nullptr;
    HANDLE searchStopEvent_ = nullptr;
};
}  // namespace quicklauncher
