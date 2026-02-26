# Search registry for DeckLink COM GUIDs
Write-Host "=== DeckLink COM Interface GUIDs ==="
$interfaces = Get-ChildItem "HKLM:\SOFTWARE\Classes\Interface" -ErrorAction SilentlyContinue
foreach ($iface in $interfaces) {
    $name = (Get-ItemProperty $iface.PSPath -ErrorAction SilentlyContinue).'(default)'
    if ($name -like "*DeckLink*") {
        Write-Host "$($iface.PSChildName) = $name"
    }
}

Write-Host "`n=== DeckLink CLSID ==="
$classes = Get-ChildItem "HKLM:\SOFTWARE\Classes\CLSID" -ErrorAction SilentlyContinue
foreach ($cls in $classes) {
    $name = (Get-ItemProperty $cls.PSPath -ErrorAction SilentlyContinue).'(default)'
    if ($name -like "*DeckLink*") {
        Write-Host "$($cls.PSChildName) = $name"
    }
}
