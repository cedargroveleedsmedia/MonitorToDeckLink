# Find and analyze DeckLink COM registration to get correct vtable info
$dll = "C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll"
if (Test-Path $dll) {
    Write-Host "Found: $dll"
    $v = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dll)
    Write-Host "Version: $($v.FileVersion)"
}

# Look for the IDL or header files installed with the SDK
$sdkPaths = @(
    "$env:ProgramFiles\Blackmagic Design\DeckLink SDK",
    "$env:ProgramFiles\Blackmagic Design\Blackmagic Desktop Video SDK",
    "C:\Program Files\Blackmagic Design",
    "C:\BMD",
    "$env:PUBLIC\Documents\Blackmagic Design"
)
foreach ($p in $sdkPaths) {
    if (Test-Path $p) {
        Write-Host "`nFound SDK dir: $p"
        Get-ChildItem $p -Recurse -Include "*.h","*.idl" -ErrorAction SilentlyContinue | Select-Object FullName
    }
}

# Also check registry for SDK install path
$regPaths = @(
    "HKLM:\SOFTWARE\Blackmagic Design\Desktop Video",
    "HKLM:\SOFTWARE\WOW6432Node\Blackmagic Design\Desktop Video"
)
foreach ($r in $regPaths) {
    if (Test-Path $r) {
        Write-Host "`nRegistry $r`:"
        Get-ItemProperty $r
    }
}
