# Restart the DeckLink driver service to release any stale device locks
Write-Host "Stopping DeckLink service..."
Stop-Service -Name "BmdService" -Force -ErrorAction SilentlyContinue
Stop-Service -Name "DeckLink" -Force -ErrorAction SilentlyContinue

# Find and kill any stale DeckLink processes
Get-Process | Where-Object { $_.Name -like "*DeckLink*" -or $_.Name -like "*Blackmagic*" } | ForEach-Object {
    Write-Host "Killing: $($_.Name)"
    Stop-Process $_ -Force
}

Start-Sleep -Seconds 2

Write-Host "Starting DeckLink service..."
Start-Service -Name "BmdService" -ErrorAction SilentlyContinue
Start-Service -Name "DeckLink" -ErrorAction SilentlyContinue

# List running DeckLink services
Get-Service | Where-Object { $_.DisplayName -like "*Blackmagic*" -or $_.DisplayName -like "*DeckLink*" }
Write-Host "Done - device locks cleared"
