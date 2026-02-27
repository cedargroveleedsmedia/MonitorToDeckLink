# Find tlbimp.exe
$tlbimp = $null
$searchPaths = @(
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\tlbimp.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\tlbimp.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.7.2 Tools\tlbimp.exe",
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.6 Tools\tlbimp.exe"
)
foreach ($p in $searchPaths) {
    if (Test-Path $p) { $tlbimp = $p; Write-Host "Found tlbimp: $p"; break }
}
if (-not $tlbimp) {
    # Search more broadly
    $tlbimp = Get-ChildItem "${env:ProgramFiles(x86)}\Microsoft SDKs" -Recurse -Filter "tlbimp.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($tlbimp) { Write-Host "Found: $tlbimp" }
}

$tlb = "C:\Users\TechOps\AppData\Local\Temp\DeckLinkAPI.tlb"
$out = "$env:TEMP\DeckLinkInterop.dll"

if ($tlbimp) {
    Write-Host "Running tlbimp..."
    & $tlbimp $tlb /out:$out /machine:X64 /namespace:DeckLinkAPI 2>&1
    if (Test-Path $out) {
        Write-Host "Generated: $out"
        # Now use ildasm to list IDeckLinkOutput methods in vtable order
        $ildasm = Get-ChildItem "${env:ProgramFiles(x86)}\Microsoft SDKs" -Recurse -Filter "ildasm.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
        if ($ildasm) {
            Write-Host "Running ildasm..."
            $il = "$env:TEMP\DeckLinkInterop.il"
            & $ildasm $out /out=$il 2>&1
            if (Test-Path $il) {
                # Extract IDeckLinkOutput interface methods in order
                $lines = Get-Content $il
                $inInterface = $false
                $methodNum = 0
                foreach ($line in $lines) {
                    if ($line -match "interface.*IDeckLinkOutput\b") { $inInterface = $true; $methodNum = 0 }
                    if ($inInterface -and $line -match "\.method") { 
                        Write-Host "  [$methodNum] $($line.Trim())"
                        $methodNum++
                    }
                    if ($inInterface -and $line -match "^} // end of class" -and $methodNum -gt 0) { $inInterface = $false }
                }
            }
        }
    }
} else {
    Write-Host "tlbimp not found - trying oleview approach"
    # Use PowerShell COM automation to read the TLB directly
    $dll = "C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll"
    
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class TlbReader {
    [DllImport("oleaut32.dll", CharSet=CharSet.Unicode)]
    public static extern int LoadTypeLibEx(string file, int regKind, out IntPtr ppTLib);
    [DllImport("oleaut32.dll")]
    public static extern int LoadRegTypeLib(ref Guid rguid, ushort wVerMajor, ushort wVerMinor, uint lcid, out IntPtr ppTLib);
}
"@
    $ptr = [IntPtr]::Zero
    $hr = [TlbReader]::LoadTypeLibEx($dll, 1, [ref]$ptr)
    Write-Host "hr=0x$($hr.ToString('X8')) ptr=$ptr"
}
