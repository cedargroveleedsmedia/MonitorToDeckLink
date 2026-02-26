@echo off
echo ============================================
echo  MonitorToDeckLink - Build Script
echo ============================================
echo.

:: Check for dotnet
where dotnet >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] .NET SDK not found.
    echo Please install from: https://dotnet.microsoft.com/download
    echo Make sure to install the .NET 6 SDK or later.
    pause
    exit /b 1
)

echo [1/3] Restoring NuGet packages...
dotnet restore MonitorToDeckLink.csproj -r win-x64
if %errorlevel% neq 0 (
    echo [ERROR] Package restore failed.
    pause
    exit /b 1
)

echo.
echo [2/3] Building and publishing (single EXE)...
dotnet publish MonitorToDeckLink.csproj ^
    /p:PublishProfile=Properties/PublishProfiles/Release-x64.pubxml
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed.
    echo.
    echo Most likely cause: DeckLink COM reference not found.
    echo Make sure Blackmagic Desktop Video is installed:
    echo   https://www.blackmagicdesign.com/support/family/capture-and-playback
    echo.
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo.
echo Output: publish\MonitorToDeckLink.exe
echo.

:: Ask to launch
set /p LAUNCH="Launch the app now? (y/n): "
if /i "%LAUNCH%"=="y" (
    start "" "publish\MonitorToDeckLink.exe"
)

pause
