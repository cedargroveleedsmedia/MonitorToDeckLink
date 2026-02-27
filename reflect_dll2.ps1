$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -like "*IDeckLinkOutput*") {
        Write-Host "=== $($t.FullName) ==="
        $i = 0
        foreach ($m in $t.GetMethods()) {
            Write-Host "  [$i] $($m.Name)"
            $i++
        }
    }
}
