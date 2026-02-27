$asm = [System.Reflection.Assembly]::LoadFile("C:\Users\TechOps\MonitorToDeckLink\DeckLinkInterop.dll")
$asm.GetTypes() | Where-Object { $_.IsInterface -and $_.Name -match "IDeckLink.*VideoFrame|IDeckLinkMutable" } | ForEach-Object {
    Write-Host "$($_.Name) = $($_.GUID)"
    $_.GetMethods() | Select-Object -First 6 | ForEach-Object { Write-Host "  $($_.Name)" }
}
