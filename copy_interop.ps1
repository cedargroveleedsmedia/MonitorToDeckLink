
$src = "$env:TEMP\DeckLinkInterop.dll"
$dst = "C:\Users\TechOps\MonitorToDeckLink\DeckLinkInterop.dll"
Copy-Item $src $dst -Force
Write-Host "Copied to $dst"
Write-Host "Size: $((Get-Item $dst).Length) bytes"

# Also check what CLSID to use for CoCreateInstance
$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)
foreach ($t in $asm.GetTypes()) {
    $attr = $t.GetCustomAttributes($false) | Where-Object { $_ -is [System.Runtime.InteropServices.GuidAttribute] }
    if ($attr -and $t.Name -notlike "*Event*") {
        Write-Host "$($t.Name): {$($attr[0].Value)}"
    }
}
