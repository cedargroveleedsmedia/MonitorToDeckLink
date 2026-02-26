@echo off
echo Checking Desktop Video version...
powershell -Command "$dll = 'C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll'; if (Test-Path $dll) { $v = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dll); Write-Host ('File version: ' + $v.FileVersion); Write-Host ('Product version: ' + $v.ProductVersion) } else { Write-Host 'DLL not found at expected path' }"
echo.
powershell -Command "$reg = 'HKLM:\SOFTWARE\Blackmagic Design\Desktop Video'; if (Test-Path $reg) { Write-Host ('Registry: ' + (Get-ItemProperty $reg).Version) } else { Write-Host 'Registry key not found' }"
pause
