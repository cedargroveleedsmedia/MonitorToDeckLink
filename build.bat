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

:: Find the DeckLink type library installed by Desktop Video
set "DECKLINK_TLB="
for /f "delims=" %%i in ('dir /s /b "C:\Program Files\Blackmagic Design\DeckLink\*.tlb" 2^>nul') do set "DECKLINK_TLB=%%i"
for /f "delims=" %%i in ('dir /s /b "C:\Program Files (x86)\Blackmagic Design\DeckLink\*.tlb" 2^>nul') do set "DECKLINK_TLB=%%i"

if not defined DECKLINK_TLB (
    echo [ERROR] DeckLink type library not found.
    echo Please install Blackmagic Desktop Video first:
    echo   https://www.blackmagicdesign.com/support/family/capture-and-playback
    pause
    exit /b 1
)
echo        Found: %DECKLINK_TLB%

:: Find tlbimp.exe - check Windows SDK and Visual Studio locations
set "TLBIMP="
for /f "delims=" %%i in ('dir /s /b "C:\Program Files (x86)\Microsoft SDKs\Windows\*\x64\tlbimp.exe" 2^>nul') do set "TLBIMP=%%i"
for /f "delims=" %%i in ('dir /s /b "C:\Program Files\Microsoft SDKs\Windows\*\x64\tlbimp.exe" 2^>nul') do set "TLBIMP=%%i"
for /f "delims=" %%i in ('dir /s /b "C:\Program Files\Microsoft Visual Studio\2022\*\bin\tlbimp.exe" 2^>nul') do set "TLBIMP=%%i"
for /f "delims=" %%i in ('dir /s /b "C:\Program Files (x86)\Microsoft Visual Studio\2019\*\bin\tlbimp.exe" 2^>nul') do set "TLBIMP=%%i"

if not defined TLBIMP (
    echo [ERROR] tlbimp.exe not found.
    echo Please install the Windows SDK: https://developer.microsoft.com/windows/downloads/windows-sdk/
    echo Or install Visual Studio Community 2022 with any workload.
    pause
    exit /b 1
)
echo        Using: %TLBIMP%

if not exist "lib" mkdir lib

"%TLBIMP%" "%DECKLINK_TLB%" /out:"lib\DeckLinkAPI.dll" /namespace:DeckLinkAPI /machine:X64
if %errorlevel% neq 0 (
    echo [ERROR] Failed to generate interop DLL.
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
