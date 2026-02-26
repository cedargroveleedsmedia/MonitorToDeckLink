# Load the DeckLink COM type library to get current interface GUIDs
$dll = "C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll"

# Try to load as type library
try {
    $tlb = [System.Runtime.InteropServices.Marshal]::GetTypeLibGuidForAssembly
    Add-Type -Path $dll -ErrorAction Stop
} catch {}

# Instead, use tlbexp or read the TLB directly from the DLL resources
# The TLB is embedded as a resource in the DLL
# Extract it with PowerShell
$bytes = [System.IO.File]::ReadAllBytes($dll)

# Search for IDeckLinkOutput GUID pattern in binary
# GUIDs we're looking for appear as raw bytes in the vtable/typelib
# Known old GUID: 1A8077F1-9FE2-4533-8147-2294305E253F
$oldGuid = [byte[]](0xF1,0x77,0x80,0x1A, 0xE2,0x9F, 0x33,0x45, 0x81,0x47, 0x22,0x94,0x30,0x5E,0x25,0x3F)
# Search for it
for ($i = 0; $i -lt $bytes.Length - 16; $i++) {
    $match = $true
    for ($j = 0; $j -lt 16; $j++) {
        if ($bytes[$i+$j] -ne $oldGuid[$j]) { $match = $false; break }
    }
    if ($match) { Write-Host "Found old IDeckLinkOutput GUID at offset $i" }
}

# Also dump all GUIDs near "IDeckLinkOutput" string
$str = [System.Text.Encoding]::Unicode.GetBytes("IDeckLinkOutput")
for ($i = 0; $i -lt $bytes.Length - $str.Length; $i++) {
    $match = $true
    for ($j = 0; $j -lt $str.Length; $j++) {
        if ($bytes[$i+$j] -ne $str[$j]) { $match = $false; break }
    }
    if ($match) {
        Write-Host "Found IDeckLinkOutput string at offset $i (0x$('{0:X}' -f $i))"
        # Show surrounding bytes as potential GUID
        $start = [Math]::Max(0, $i-20)
        $end = [Math]::Min($bytes.Length, $i+$str.Length+20)
        $hex = ($bytes[$start..$end] | ForEach-Object { '{0:X2}' -f $_ }) -join ' '
        Write-Host "  Context: $hex"
    }
}
Write-Host "Done searching"
