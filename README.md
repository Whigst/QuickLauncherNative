# QuickLauncher Native

A native Win32 launcher focused on low background overhead.

## Features

- Custom global hotkey summon/hide, default `Alt+Space`
- Compact window layout focused on query and results
- Debounced background search keeps typing smooth on large indexes
- Search uses token-prefix candidate buckets plus previous-match narrowing, so longer queries get cheaper instead of rescanning everything
- Pre-index fixed drives and cache launchable files
- Search and open folders themselves
- Search `.exe`, `.bat`, `.cmd`, `.ps1`, `.vbs`, `.lnk`, `.msc`, and similar launch targets
- Fuzzy matching plus initials matching for acronym and pinyin-initial style search
- Frequent launches and query-specific history move results upward
- Launchable results like `.lnk`, `.exe`, `.bat`, and `.cmd` are ranked ahead of folder matches
- Empty query uses a cached home list, so hotkey summon can show recommendations without a fresh full-index sort
- Start Menu, Desktop, and pinned shortcut entries get extra ranking priority
- Same-name shortcut-heavy results are lightly deduplicated to reduce noise
- Type an existing full path or URL to launch it directly
- Result list uses native Windows file icons
- Result list shows a compact source directory tag
- Automatically debounces changes in quick-access folders and configured roots, then refreshes in the background
- Hotkey summon clears the query, `Ctrl+A` selects all input, `Ctrl+Shift+Enter` runs as administrator, `Shift+Enter` reveals in Explorer, and `Ctrl+C` copies the full path
- Right-click a result, or press the menu key / `Shift+F10`, for launch, admin, reveal, and copy actions
- Main window includes a `Run at Startup` toggle switch, while the tray menu keeps the same option as a backup path
- Tray menu with open, rebuild index, settings, autostart toggle, and exit
- Folder entries are indexed from your user profile plus configured `include` and `priority` directories to keep background memory lower

## Settings

- Settings file: `%LOCALAPPDATA%\\QuickLauncherNative\\settings.ini`
- Supported lines:
  - `hotkey=Alt+Space`
  - `include=E:\\Tools`
  - `exclude=E:\\SteamLibrary\\steamapps`
  - `priority=E:\\Tools`
- After editing the file, use the tray menu `Reload Settings` to apply and rebuild.

## Autostart

- Use the tray menu `Run at Startup` to add or remove the launcher from `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run`.

## Build

```bat
build.bat
```

## Output

- `QuickLauncherNative.exe`
- The first run after this update rebuilds the cache into `launcher-index-v3.tsv`
