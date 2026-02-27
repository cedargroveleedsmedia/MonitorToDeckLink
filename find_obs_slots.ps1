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
