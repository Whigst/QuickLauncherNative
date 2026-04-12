# QuickLauncher Native / QuickLauncher 原生版

A native Win32 launcher focused on low background overhead. / 一个专注于低后台占用的原生 Win32 快捷启动器。

## Features / 功能

- Custom global hotkey summon or hide, default `Alt+Space` / 支持自定义全局热键呼出或隐藏，默认是 `Alt+Space`
- Compact window layout focused on query and results / 紧凑窗口布局，只保留搜索和结果区域
- Debounced background search keeps typing smooth on large indexes / 后台搜索带防抖，在大索引下输入更平滑
- Search uses token-prefix candidate buckets plus previous-match narrowing, so longer queries get cheaper instead of rescanning everything / 搜索使用分词前缀候选桶和上次匹配结果收窄，输入越长，检索越便宜，不会每次全量重扫
- Pre-index fixed drives and cache launchable files / 预扫描固定磁盘并缓存可启动文件
- Search and open folders themselves / 支持搜索并直接打开文件夹
- Search `.exe`, `.bat`, `.cmd`, `.ps1`, `.vbs`, `.lnk`, `.msc`, and similar launch targets / 支持搜索 `.exe`、`.bat`、`.cmd`、`.ps1`、`.vbs`、`.lnk`、`.msc` 等启动目标
- Fuzzy matching plus initials matching for acronym and pinyin-initial style search / 支持模糊匹配和首字母匹配，适合缩写和拼音首字母搜索
- Frequent launches and query-specific history move results upward / 常用结果和查询历史会提升排序
- Prefix search with `/.bat:` or `/.lnk:` to prioritize that extension, for example `/.bat: deploy` / 可以用 `/.bat:` 或 `/.lnk:` 这样的前缀优先显示对应后缀，例如 `/.bat: deploy`
- Prefix search with `/dir:` to prioritize directories first and browse the current path level, for example `/dir:E:\\`, `/dir:E:\\Tools`, or `/dir:\"E:\\Program Files\" obsidian` / 可以用 `/dir:` 前缀优先显示目录，并浏览当前路径层级，例如 `/dir:E:\\`、`/dir:E:\\Tools` 或 `/dir:\"E:\\Program Files\" obsidian`
- Launchable results like `.lnk`, `.exe`, `.bat`, and `.cmd` are ranked ahead of folder matches / `.lnk`、`.exe`、`.bat`、`.cmd` 这类可启动结果会排在文件夹之前
- Empty query uses a cached home list, so hotkey summon can show recommendations without a fresh full-index sort / 空查询会使用缓存首页列表，热键呼出时不用重新全量排序也能显示推荐
- Start Menu, Desktop, and pinned shortcut entries get extra ranking priority / 开始菜单、桌面和固定入口会获得额外排序加成
- Same-name shortcut-heavy results are lightly deduplicated to reduce noise / 对同名快捷方式较多的结果做轻量去重，减少噪声
- Type an existing full path or URL to launch it directly / 输入现有完整路径或网址可直接启动
- Result list uses native Windows file icons / 结果列表使用原生 Windows 文件图标
- Result list shows a compact source directory tag / 结果列表会显示紧凑的来源目录标签
- Automatically debounces changes in quick-access folders and configured roots, then refreshes in the background / 会对常用目录和配置目录的变化做防抖，然后后台自动刷新
- Hotkey summon clears the query, `Ctrl+A` selects all input, `Ctrl+Shift+Enter` runs as administrator, `Shift+Enter` reveals in Explorer, and `Ctrl+C` copies the full path / 热键呼出会清空输入，`Ctrl+A` 全选输入，`Ctrl+Shift+Enter` 以管理员运行，`Shift+Enter` 在资源管理器中定位，`Ctrl+C` 复制完整路径
- Right-click a result, or press the menu key or `Shift+F10`, for launch, admin, reveal, and copy actions / 右键结果，或按菜单键或 `Shift+F10`，可执行启动、管理员运行、打开位置和复制路径
- Main window includes a `Run at Startup` toggle switch, while the tray menu keeps the same option as a backup path / 主窗口内置 `Run at Startup` 开关，同时托盘菜单保留相同选项作为备用入口
- Tray menu with open, rebuild index, settings, autostart toggle, and exit / 托盘菜单提供打开、重建索引、设置、开机自启开关和退出
- Folder entries are indexed from your user profile plus configured `include` and `priority` directories to keep background memory lower / 文件夹索引默认来自用户目录以及配置中的 `include` 和 `priority` 目录，以降低后台内存占用

## Settings / 设置

- Settings file: `%LOCALAPPDATA%\\QuickLauncherNative\\settings.ini` / 设置文件：`%LOCALAPPDATA%\\QuickLauncherNative\\settings.ini`
- Supported lines / 支持的配置行：
  - `hotkey=Alt+Space` / 设置热键
  - `include=E:\\Tools` / 额外纳入扫描的目录
  - `exclude=E:\\SteamLibrary\\steamapps` / 排除扫描的目录
  - `priority=E:\\Tools` / 排序优先目录
- After editing the file, use the tray menu `Reload Settings` to apply and rebuild. / 修改文件后，可通过托盘菜单 `Reload Settings` 重新加载并重建索引。

## Autostart / 开机自启

- Use the tray menu or the main-window toggle `Run at Startup` to add or remove the launcher from `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run`. / 可通过托盘菜单或主窗口中的 `Run at Startup` 开关，将启动器加入或移出 `HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run`。

## Build / 构建

```bat
build.bat
```

## Output / 输出

- `QuickLauncherNative.exe` / 主程序文件 `QuickLauncherNative.exe`
- The first run after this update rebuilds the cache into `launcher-index-v3.tsv` / 本次更新后的首次运行会把缓存重建到 `launcher-index-v3.tsv`
