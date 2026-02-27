$il = "$env:TEMP\DeckLinkInterop.il"
Write-Host "File size: $((Get-Item $il -ErrorAction SilentlyContinue).Length) bytes"
Write-Host "First 100 lines:"
Get-Content $il -TotalCount 100

Write-Host "`n--- Searching for DeckLink ---"
Select-String -Path $il -Pattern "DeckLink" | Select-Object -First 30 | ForEach-Object { Write-Host $_.Line }
