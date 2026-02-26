@echo off
echo Searching for DeckLink SDK headers...
dir /s /b "C:\Program Files\Blackmagic Design\*.h" 2>nul
dir /s /b "C:\Program Files (x86)\Blackmagic Design\*.h" 2>nul
echo.
echo If no headers found, extracting IDeckLinkOutput vtable from DLL using powershell...
powershell -Command "& {
    $dll = 'C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll'
    if (Test-Path $dll) {
        Write-Host 'DLL found, checking exports...'
        $bytes = [System.IO.File]::ReadAllBytes($dll)
        Write-Host ('DLL size: ' + $bytes.Length + ' bytes')
    }
}"
echo.
echo Checking Desktop Video version...
powershell -Command "& {
    $dll = 'C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll'
    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dll)
    Write-Host ('Version: ' + $vi.FileVersion)
    Write-Host ('Product: ' + $vi.ProductVersion)
}"
pause
