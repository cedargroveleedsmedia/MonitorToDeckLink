$il = "$env:TEMP\DeckLinkInterop.il"
if (-not (Test-Path $il)) { Write-Host "IL file not found at $il"; exit }

$lines = Get-Content $il
$inInterface = $false
$methodNum = 0
$depth = 0

foreach ($line in $lines) {
    if ($line -match "\.class.*interface.*IDeckLinkOutput\b") { 
        $inInterface = $true; $methodNum = 0; $depth = 0
        Write-Host "=== IDeckLinkOutput ==="
    }
    if ($inInterface) {
        if ($line -match "\{") { $depth++ }
        if ($line -match "\}") { $depth--; if ($depth -le 0) { $inInterface = $false } }
        if ($line -match "\.method.*instance.*\s(\w+)\s*\(") {
            Write-Host "  [$methodNum] $($Matches[1])"
            $methodNum++
        }
    }
}
