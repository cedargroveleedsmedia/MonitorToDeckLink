$dll = "$env:TEMP\DeckLinkInterop.dll"
Write-Host "Loading $dll ($(((Get-Item $dll).Length)) bytes)"
$asm = [System.Reflection.Assembly]::LoadFile($dll)
$types = $asm.GetTypes()
Write-Host "Types found: $($types.Count)"
foreach ($t in $types) {
    if ($t.Name -like "*IDeckLinkOutput*") {
        Write-Host "`n=== $($t.FullName) ==="
        $methods = $t.GetMethods()
        $i = 0
        foreach ($m in $methods) {
            Write-Host "  [$i] $($m.Name)($($m.GetParameters() | ForEach-Object { $_.ParameterType.Name + ' ' + $_.Name } | Join-String -Separator ', '))"
            $i++
        }
    }
}
