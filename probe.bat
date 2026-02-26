@echo off
echo Checking Desktop Video version...
powershell -Command "& {
    $paths = @(
        'C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll',
        'C:\Program Files (x86)\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll'
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            $v = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($p)
            Write-Host ('DLL: ' + $p)
            Write-Host ('File version: ' + $v.FileVersion)
            Write-Host ('Product version: ' + $v.ProductVersion)
        }
    }
    # Also check registry
    $reg = 'HKLM:\SOFTWARE\Blackmagic Design\Desktop Video'
    if (Test-Path $reg) {
        $rv = Get-ItemProperty $reg
        Write-Host ('Registry version: ' + $rv.Version)
    }
}"
pause
