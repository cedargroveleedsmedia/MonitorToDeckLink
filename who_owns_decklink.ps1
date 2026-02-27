# Find processes with handles to DeckLink DLL
$dll = "DeckLinkAPI64.dll"
Write-Host "Processes using $dll :"
Get-Process | ForEach-Object {
    try {
        $modules = $_.Modules | Where-Object { $_.ModuleName -like "*DeckLink*" -or $_.ModuleName -like "*Blackmagic*" }
        if ($modules) {
            Write-Host "  PID=$($_.Id) Name=$($_.Name)"
            $modules | ForEach-Object { Write-Host "    -> $($_.FileName)" }
        }
    } catch {}
}
