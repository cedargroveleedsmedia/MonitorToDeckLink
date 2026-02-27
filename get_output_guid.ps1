$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -like "*IDeckLinkOutput*") {
        $guid = $t.GUID
        Write-Host "$($t.Name) GUID: {$guid}"
    }
}

# Also copy the interop DLL to project directory
$src = "$env:TEMP\DeckLinkInterop.dll"
$dst = "C:\Users\TechOps\MonitorToDeckLink\DeckLinkInterop.dll"
Copy-Item $src $dst -Force
Write-Host "`nCopied interop DLL to project: $dst"
