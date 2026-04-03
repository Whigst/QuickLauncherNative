@echo off
setlocal
g++ -std=c++20 -Os -s -municode -mwindows -finput-charset=UTF-8 -fexec-charset=UTF-8 LauncherApp.cpp -o QuickLauncherNative.exe -lcomctl32 -lshell32 -lole32 -luuid -lgdi32
if errorlevel 1 exit /b %errorlevel%
echo Built QuickLauncherNative.exe
