#pragma once

#include <windows.h>

#include <algorithm>
#include <cwctype>
#include <sstream>
#include <string>
#include <unordered_map>
#include <vector>

namespace quicklauncher
{
constexpr long long kTicksPerSecond = 10000000LL;
constexpr long long kTicksPerHour = kTicksPerSecond * 60LL * 60LL;
constexpr long long kTicksPerDay = kTicksPerHour * 24LL;

struct LaunchEntry
{
    std::wstring name;
    std::wstring extension;
    std::wstring fullPath;
    std::wstring sourceTag;
    std::wstring searchNameCompact;
    std::wstring searchPathCompact;
    std::wstring searchInitials;
    bool isDirectory = false;
    bool isUrl = false;
    int iconIndex = -1;
    long long lastWriteUtcTicks = 0;
};

struct UsageStat
{
    int launchCount = 0;
    long long lastLaunchUtcTicks = 0;
};

struct QueryBucket
{
    long long lastLaunchUtcTicks = 0;
    std::unordered_map<std::wstring, int> launchCounts;
};

struct SearchResult
{
    size_t entryIndex = 0;
    int score = 0;
    int usageCount = 0;
};

inline std::wstring ToLowerCopy(const std::wstring& value)
{
    std::wstring result = value;
    std::transform(result.begin(), result.end(), result.begin(), [](wchar_t ch)
    {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return result;
}

inline bool IsCjkCharacter(wchar_t ch)
{
    return (ch >= L'\u3400' && ch <= L'\u4dbf')
        || (ch >= L'\u4e00' && ch <= L'\u9fff')
        || (ch >= L'\uf900' && ch <= L'\ufaff');
}

inline wchar_t GetPinyinInitial(wchar_t ch)
{
    static constexpr std::pair<int, wchar_t> kPinyinInitialRanges[] =
    {
        {45217, L'a'},
        {45253, L'b'},
        {45761, L'c'},
        {46318, L'd'},
        {46826, L'e'},
        {47010, L'f'},
        {47297, L'g'},
        {47614, L'h'},
        {48119, L'j'},
        {49062, L'k'},
        {49324, L'l'},
        {49896, L'm'},
        {50371, L'n'},
        {50614, L'o'},
        {50622, L'p'},
        {50906, L'q'},
        {51387, L'r'},
        {51446, L's'},
        {52218, L't'},
        {52698, L'w'},
        {52980, L'x'},
        {53689, L'y'},
        {54481, L'z'}
    };

    char bytes[4] = {};
    const int count = WideCharToMultiByte(936, 0, &ch, 1, bytes, 4, nullptr, nullptr);
    if (count < 2)
    {
        return static_cast<wchar_t>(std::towlower(ch));
    }

    const int code = (static_cast<unsigned char>(bytes[0]) << 8) + static_cast<unsigned char>(bytes[1]);
    for (int index = static_cast<int>(std::size(kPinyinInitialRanges)) - 1; index >= 0; --index)
    {
        if (code >= kPinyinInitialRanges[index].first)
        {
            return kPinyinInitialRanges[index].second;
        }
    }

    return static_cast<wchar_t>(std::towlower(ch));
}

inline std::wstring NormalizeCompact(const std::wstring& value)
{
    std::wstring result;
    result.reserve(value.size());

    for (const auto ch : value)
    {
        if (std::iswalnum(ch) || IsCjkCharacter(ch))
        {
            result.push_back(static_cast<wchar_t>(std::towlower(ch)));
        }
    }

    return result;
}

inline std::wstring BuildInitials(const std::wstring& value)
{
    std::wstring result;
    result.reserve(value.size());
    bool startOfToken = true;

    for (const auto ch : value)
    {
        if (std::iswalnum(ch))
        {
            if (startOfToken)
            {
                result.push_back(static_cast<wchar_t>(std::towlower(ch)));
            }

            startOfToken = false;
            continue;
        }

        if (IsCjkCharacter(ch))
        {
            result.push_back(GetPinyinInitial(ch));
            startOfToken = true;
            continue;
        }

        startOfToken = true;
    }

    return result;
}

inline bool StartsWith(const std::wstring& value, const std::wstring& prefix)
{
    return value.size() >= prefix.size()
        && std::equal(prefix.begin(), prefix.end(), value.begin());
}

inline int SubsequenceScore(const std::wstring& pattern, const std::wstring& candidate)
{
    if (pattern.empty() || pattern.size() > candidate.size())
    {
        return 0;
    }

    size_t candidateIndex = 0;
    size_t matchedCharacters = 0;
    int firstMatchIndex = -1;
    int lastMatchIndex = -1;
    int consecutiveRun = 0;
    int score = 0;

    for (const auto patternChar : pattern)
    {
        bool found = false;
        while (candidateIndex < candidate.size())
        {
            const auto candidateChar = candidate[candidateIndex];
            if (candidateChar == patternChar)
            {
                if (firstMatchIndex < 0)
                {
                    firstMatchIndex = static_cast<int>(candidateIndex);
                }

                consecutiveRun = lastMatchIndex >= 0 && static_cast<int>(candidateIndex) == lastMatchIndex + 1
                    ? consecutiveRun + 1
                    : 1;

                score += 120 + std::min(consecutiveRun * 26, 180);
                lastMatchIndex = static_cast<int>(candidateIndex);
                ++candidateIndex;
                ++matchedCharacters;
                found = true;
                break;
            }

            consecutiveRun = 0;
            ++candidateIndex;
        }

        if (!found)
        {
            return 0;
        }
    }

    score += std::max(0, 180 - std::max(firstMatchIndex, 0) * 7);
    score += std::max(0, 90 - static_cast<int>(candidate.size() - matchedCharacters) * 4);
    return score;
}

inline int ScorePattern(const std::wstring& queryCompact, const std::wstring& candidateCompact)
{
    if (queryCompact.empty() || candidateCompact.empty())
    {
        return 0;
    }

    if (queryCompact == candidateCompact)
    {
        return 3200;
    }

    if (StartsWith(candidateCompact, queryCompact))
    {
        return 2500 - std::min(static_cast<int>(candidateCompact.size()), 120);
    }

    const auto containsIndex = candidateCompact.find(queryCompact);
    if (containsIndex != std::wstring::npos)
    {
        return 1900 - std::min(static_cast<int>(containsIndex) * 22, 500);
    }

    return SubsequenceScore(queryCompact, candidateCompact);
}

inline long long CurrentUtcTicks()
{
    FILETIME fileTime{};
    GetSystemTimeAsFileTime(&fileTime);
    ULARGE_INTEGER value{};
    value.LowPart = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return static_cast<long long>(value.QuadPart);
}

inline long long FileTimeToTicks(const FILETIME& fileTime)
{
    ULARGE_INTEGER value{};
    value.LowPart = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return static_cast<long long>(value.QuadPart);
}

inline std::wstring TicksToLocalTimeString(long long utcTicks)
{
    if (utcTicks <= 0)
    {
        return L"Not indexed yet";
    }

    FILETIME utcFileTime{};
    ULARGE_INTEGER value{};
    value.QuadPart = static_cast<unsigned long long>(utcTicks);
    utcFileTime.dwLowDateTime = value.LowPart;
    utcFileTime.dwHighDateTime = value.HighPart;

    FILETIME localFileTime{};
    SYSTEMTIME localSystemTime{};
    FileTimeToLocalFileTime(&utcFileTime, &localFileTime);
    FileTimeToSystemTime(&localFileTime, &localSystemTime);

    wchar_t buffer[64] = {};
    swprintf_s(
        buffer,
        std::size(buffer),
        L"%04u-%02u-%02u %02u:%02u",
        localSystemTime.wYear,
        localSystemTime.wMonth,
        localSystemTime.wDay,
        localSystemTime.wHour,
        localSystemTime.wMinute);
    return buffer;
}

inline int CalculateRecencyBoost(long long lastLaunchUtcTicks)
{
    if (lastLaunchUtcTicks <= 0)
    {
        return 0;
    }

    const auto age = CurrentUtcTicks() - lastLaunchUtcTicks;
    if (age <= 12LL * kTicksPerHour)
    {
        return 520;
    }

    if (age <= 3LL * kTicksPerDay)
    {
        return 360;
    }

    if (age <= 14LL * kTicksPerDay)
    {
        return 200;
    }

    if (age <= 45LL * kTicksPerDay)
    {
        return 80;
    }

    return 0;
}

inline int ExtensionBoost(const std::wstring& extension)
{
    if (extension == L".exe")
    {
        return 160;
    }

    if (extension == L".lnk")
    {
        return 120;
    }

    if (extension == L".bat" || extension == L".cmd")
    {
        return 90;
    }

    if (extension == L".ps1" || extension == L".vbs" || extension == L".js")
    {
        return 80;
    }

    if (extension == L".msc" || extension == L".appref-ms")
    {
        return 60;
    }

    return 40;
}

inline std::wstring EscapeField(const std::wstring& value)
{
    std::wstring result;
    result.reserve(value.size());

    for (const auto ch : value)
    {
        switch (ch)
        {
        case L'\\':
            result += L"\\\\";
            break;
        case L'\t':
            result += L"\\t";
            break;
        case L'\n':
            result += L"\\n";
            break;
        case L'\r':
            result += L"\\r";
            break;
        default:
            result.push_back(ch);
            break;
        }
    }

    return result;
}

inline std::vector<std::wstring> SplitEscapedLine(const std::wstring& line)
{
    std::vector<std::wstring> fields;
    std::wstring current;
    bool escape = false;

    for (const auto ch : line)
    {
        if (escape)
        {
            switch (ch)
            {
            case L't':
                current.push_back(L'\t');
                break;
            case L'n':
                current.push_back(L'\n');
                break;
            case L'r':
                current.push_back(L'\r');
                break;
            case L'\\':
                current.push_back(L'\\');
                break;
            default:
                current.push_back(ch);
                break;
            }

            escape = false;
            continue;
        }

        if (ch == L'\\')
        {
            escape = true;
            continue;
        }

        if (ch == L'\t')
        {
            fields.push_back(current);
            current.clear();
            continue;
        }

        current.push_back(ch);
    }

    if (escape)
    {
        current.push_back(L'\\');
    }

    fields.push_back(current);
    return fields;
}

inline std::wstring JoinPath(const std::wstring& left, const std::wstring& right)
{
    if (left.empty())
    {
        return right;
    }

    if (left.back() == L'\\')
    {
        return left + right;
    }

    return left + L'\\' + right;
}

inline bool IsLaunchableExtension(const std::wstring& extension)
{
    static const std::vector<std::wstring> kExtensions =
    {
        L".exe", L".bat", L".cmd", L".com", L".ps1", L".vbs", L".vbe", L".js",
        L".jse", L".wsf", L".msc", L".lnk", L".msi", L".appref-ms"
    };

    const auto lower = ToLowerCopy(extension);
    return std::find(kExtensions.begin(), kExtensions.end(), lower) != kExtensions.end();
}

inline bool ShouldSkipDirectory(const std::wstring& path)
{
    static const std::vector<std::wstring> kMarkers =
    {
        L"\\$recycle.bin\\",
        L"\\system volume information\\",
        L"\\windows\\winsxs\\",
        L"\\windows\\installer\\",
        L"\\program files\\windowsapps\\",
        L"\\programdata\\package cache\\"
    };

    const auto lower = ToLowerCopy(path);
    return std::any_of(kMarkers.begin(), kMarkers.end(), [&](const std::wstring& marker)
    {
        return lower.find(marker) != std::wstring::npos;
    });
}

inline std::wstring GetDirectoryPart(const std::wstring& path)
{
    const auto slashIndex = path.find_last_of(L"\\/");
    if (slashIndex == std::wstring::npos)
    {
        return L"";
    }

    return path.substr(0, slashIndex);
}

inline std::wstring GetLeafName(const std::wstring& path)
{
    const auto slashIndex = path.find_last_of(L"\\/");
    if (slashIndex == std::wstring::npos)
    {
        return path;
    }

    if (slashIndex + 1 >= path.size())
    {
        return path.substr(0, slashIndex);
    }

    return path.substr(slashIndex + 1);
}

inline std::wstring GetFileNameWithoutExtension(const std::wstring& fileName)
{
    const auto dotIndex = fileName.find_last_of(L'.');
    if (dotIndex == std::wstring::npos)
    {
        return fileName;
    }

    return fileName.substr(0, dotIndex);
}

inline std::wstring GetExtension(const std::wstring& fileName)
{
    const auto dotIndex = fileName.find_last_of(L'.');
    if (dotIndex == std::wstring::npos)
    {
        return L"";
    }

    return ToLowerCopy(fileName.substr(dotIndex));
}

inline std::vector<std::wstring> TokenizeQuery(const std::wstring& query)
{
    std::vector<std::wstring> tokens;
    std::wstring current;

    for (const auto ch : query)
    {
        if (ch == L' ' || ch == L'\t' || ch == L'-' || ch == L'_' || ch == L'.' || ch == L'\\' || ch == L'/')
        {
            if (!current.empty())
            {
                tokens.push_back(NormalizeCompact(current));
                current.clear();
            }

            continue;
        }

        current.push_back(ch);
    }

    if (!current.empty())
    {
        tokens.push_back(NormalizeCompact(current));
    }

    tokens.erase(
        std::remove_if(tokens.begin(), tokens.end(), [](const std::wstring& token)
        {
            return token.empty();
        }),
        tokens.end());
    std::sort(tokens.begin(), tokens.end());
    tokens.erase(std::unique(tokens.begin(), tokens.end()), tokens.end());
    return tokens;
}
}  // namespace quicklauncher
