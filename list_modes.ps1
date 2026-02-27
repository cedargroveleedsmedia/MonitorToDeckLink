# Use the interop DLL to enumerate actual supported modes from the device
$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)

# Show all BMDDisplayMode enum values
$modeEnum = $asm.GetTypes() | Where-Object { $_.Name -eq "_BMDDisplayMode" }
if ($modeEnum) {
    Write-Host "=== _BMDDisplayMode values ==="
    [System.Enum]::GetValues($modeEnum) | ForEach-Object {
        $val = [int]$_
        Write-Host "  $_ = 0x$($val.ToString('X8'))"
    }
}
