$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -like "*IDeckLinkMutableVideoFrame*" -or $t.Name -like "*IDeckLinkVideoBuffer*" -or ($t.Name -like "*IDeckLinkVideoFrame*" -and $t.Name -notlike "*3D*" -and $t.Name -notlike "*Input*" -and $t.Name -notlike "*Metadata*")) {
        Write-Host "`n=== $($t.Name) ==="
        $i = 0
        foreach ($m in $t.GetMethods()) { Write-Host "  [$i] $($m.Name)"; $i++ }
    }
}
