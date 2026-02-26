@echo off
echo Inspecting DeckLink interop DLL methods on IDeckLinkVideoFrame...
echo.

:: Use ildasm if available, otherwise powershell reflection
set "ILDASM="
for /f "delims=" %%i in ('dir /s /b "C:\Program Files (x86)\Microsoft SDKs\Windows\*\x86\ildasm.exe" 2^>nul') do set "ILDASM=%%i"
for /f "delims=" %%i in ('dir /s /b "C:\Program Files\Microsoft SDKs\Windows\*\x86\ildasm.exe" 2^>nul') do set "ILDASM=%%i"

if defined ILDASM (
    echo Using ildasm: %ILDASM%
    "%ILDASM%" /text /pubonly "lib\DeckLinkAPI.dll" | findstr /i "GetBytes\|GetRowBytes\|IDeckLinkMutableVideo\|IDeckLinkVideo"
) else (
    echo ildasm not found, using PowerShell reflection...
    powershell -Command "& { $a = [System.Reflection.Assembly]::LoadFile('%cd%\lib\DeckLinkAPI.dll'); $t = $a.GetTypes() | Where-Object { $_.Name -like '*DeckLinkMutableVideo*' -or $_.Name -like '*DeckLinkVideo*' }; foreach ($type in $t) { Write-Host '--- ' $type.FullName; $type.GetMethods() | ForEach-Object { Write-Host '  ' $_.Name } } }"
)
echo.
pause
