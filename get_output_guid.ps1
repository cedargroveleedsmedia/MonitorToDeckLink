$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -like "*IDeckLinkOutput*") {
        $guid = $t.GUID
        Write-Host "$($t.Name) GUID: {$guid}"
    }
}
