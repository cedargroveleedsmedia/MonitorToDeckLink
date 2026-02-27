# Test ScheduleVideoFrame E_INVALIDARG by checking supported modes
# Load the interop DLL to use proper COM instead of vtable hacks
$asm = [System.Reflection.Assembly]::LoadFile("C:\Users\TechOps\MonitorToDeckLink\DeckLinkInterop.dll")

Write-Host "Testing via DeckLinkInterop.dll..."
Write-Host ""

# Create iterator
$iterType = $asm.GetType("DeckLinkLib.IDeckLinkIterator")
$iterCoclassType = $asm.GetTypes() | Where-Object { $_.Name -eq "CDeckLinkIterator" }
Write-Host "CDeckLinkIterator type: $iterCoclassType"

$deckLinkIterType = $asm.GetTypes() | Where-Object { $_.Name -match "Iterator" } | Select-Object Name
$deckLinkIterType | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "All exported types:"
$asm.GetTypes() | Where-Object { $_.IsPublic } | Select-Object Name | ForEach-Object { Write-Host "  $($_.Name)" }
