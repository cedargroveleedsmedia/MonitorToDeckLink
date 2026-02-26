@echo off
echo ============================================
echo  MonitorToDeckLink - Build Script
echo ============================================
echo.

:: ── Step 1: Check dotnet ─────────────────────────────────────────────────────
where dotnet >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] .NET SDK not found.
    echo Please install from: https://dotnet.microsoft.com/download/dotnet/6.0
    pause
    exit /b 1
)

:: ── Step 2: Generate DeckLink interop DLL if not already present ─────────────
if exist "lib\DeckLinkAPI.dll" (
    echo [OK] DeckLink interop DLL already exists, skipping generation.
    goto :build
)

echo [1/4] Generating DeckLink interop DLL...

:: DeckLinkAPI64.dll is the COM server
set "DECKLINK_DLL=C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll"
if not exist "%DECKLINK_DLL%" (
    echo [ERROR] DeckLinkAPI64.dll not found at:
    echo         %DECKLINK_DLL%
    pause
    exit /b 1
)
echo        Found: %DECKLINK_DLL%

:: Hardcoded known location (with spaces in path)
set "TLBIMP=C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.2 Tools\x64\TlbImp.exe"

if not exist "%TLBIMP%" (
    :: Fallback: search dynamically
    set "TLBIMP="
    for /f "delims=" %%i in ('dir /s /b "C:\Program Files (x86)\Microsoft SDKs\Windows\*.exe" 2^>nul ^| findstr /i "tlbimp"') do set "TLBIMP=%%i"
)

if not exist "%TLBIMP%" (
    echo [ERROR] TlbImp.exe not found.
    pause
    exit /b 1
)
echo        Using: %TLBIMP%

if not exist "lib" mkdir lib

"%TLBIMP%" "%DECKLINK_DLL%" /out:"lib\DeckLinkAPI.dll" /namespace:DeckLinkAPI /machine:X64
if %errorlevel% neq 0 (
    echo [ERROR] Could not generate interop DLL. Try running as Administrator.
    pause
    exit /b 1
)
echo        Generated: lib\DeckLinkAPI.dll

:build
:: ── Step 3: Restore packages ─────────────────────────────────────────────────
echo.
echo [2/4] Restoring NuGet packages...
dotnet restore MonitorToDeckLink.csproj -r win-x64
if %errorlevel% neq 0 (
    echo [ERROR] Package restore failed.
    pause
    exit /b 1
)

:: ── Step 4: Publish ──────────────────────────────────────────────────────────
echo.
echo [3/4] Building and publishing single EXE...
dotnet publish MonitorToDeckLink.csproj /p:PublishProfile=Properties/PublishProfiles/Release-x64.pubxml
if %errorlevel% neq 0 (
    echo [ERROR] Build failed. See errors above.
    pause
    exit /b 1
)

echo.
echo [4/4] Done!
echo.
echo Output: publish\MonitorToDeckLink.exe
echo.

set /p LAUNCH="Launch the app now? (y/n): "
if /i "%LAUNCH%"=="y" start "" "publish\MonitorToDeckLink.exe"

pause
