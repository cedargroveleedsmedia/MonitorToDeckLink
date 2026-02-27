$dll = "$env:TEMP\DeckLinkInterop.dll"
$asm = [System.Reflection.Assembly]::LoadFile($dll)

# Get the IDeckLinkOutput type (current = 1A8077F1)
$outType = $asm.GetTypes() | Where-Object { $_.Name -eq "IDeckLinkOutput" }
Write-Host "IDeckLinkOutput GUID: $($outType.GUID)"
Write-Host "Methods:"
$i = 0
foreach ($m in $outType.GetMethods()) {
    Write-Host "  [$i] $($m.Name)  params=($( ($m.GetParameters() | ForEach-Object { $_.ParameterType.Name }) -join ', '))"
    $i++
}

# Also check IDeckLinkVideoFrame and IDeckLinkMutableVideoFrame
Write-Host ""
Write-Host "=== All interfaces ==="
$tlb = [System.Runtime.InteropServices.TypeLibConverter]::new()
$asm = [System.Reflection.Assembly]::LoadFile("C:\Users\TechOps\MonitorToDeckLink\DeckLinkInterop.dll")
$asm.GetTypes() | Where-Object { $_.IsInterface -and $_.Name -match "IDeckLinkMutableVideoFrame|IDeckLinkVideoFrame$" } | ForEach-Object {
    Write-Host ""
    Write-Host "Interface: $($_.Name)  GUID: $($_.GUID)"
    $_.GetMethods() | ForEach-Object {
        $params = ($_.GetParameters() | ForEach-Object { $_.ParameterType.Name }) -join ", "
        Write-Host "  $($_.Name)  ($params)"
    }
}
