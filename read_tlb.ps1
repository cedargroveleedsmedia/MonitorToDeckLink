# Load the type library from DeckLinkAPI64.dll and list IDeckLinkOutput methods in order
$dll = "C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll"

# Use COM to load the type library
$tlb = $null
try {
    $tlb = [System.Runtime.InteropServices.Marshal]::LoadTypeLibEx($dll, 0)
    Write-Host "Loaded type library"
} catch {
    Write-Host "LoadTypeLibEx failed: $_"
}

if ($tlb -eq $null) {
    # Try the registered TLB
    try {
        $tlb = [System.Runtime.InteropServices.Marshal]::LoadTypeLib($dll)
        Write-Host "Loaded via LoadTypeLib"
    } catch {
        Write-Host "LoadTypeLib also failed: $_"
    }
}

# Alternative: use tlbinf32 or oleaut32 directly
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00020402-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ITypeLib {
    void GetTypeInfoCount();
    void GetTypeInfo(int index, out IntPtr ppTI);
    void GetTypeInfoType(int index, out int pTKind);
    void GetTypeInfoOfGuid(ref Guid guid, out IntPtr ppTInfo);
    void GetLibAttr(out IntPtr ppTLibAttr);
    void GetTypeComp(out IntPtr ppTComp);
    void GetDocumentation(int index, out string pbstrName, out string pbstrDocString, out int pdwHelpContext, out string pbstrHelpFile);
    void IsName(string szNameBuf, int lHashVal, out bool pfName);
    void FindName(string szNameBuf, int lHashVal, IntPtr[] ppTInfo, int[] rgMemId, ref short pcFound);
    void ReleaseTLibAttr(IntPtr pTLibAttr);
}

public class TlbLoader {
    [DllImport("oleaut32.dll", CharSet=CharSet.Unicode)]
    public static extern int LoadTypeLibEx(string szFile, int regKind, out IntPtr pptlib);
}
"@

$pptlib = [IntPtr]::Zero
$hr = [TlbLoader]::LoadTypeLibEx($dll, 1, [ref]$pptlib)
Write-Host "LoadTypeLibEx hr=0x$($hr.ToString('X8')) ptr=0x$($pptlib.ToString('X'))"

if ($hr -eq 0 -and $pptlib -ne [IntPtr]::Zero) {
    Write-Host "Type library loaded - now use ildasm or tlbdump to inspect"
    # Try generating interop
    $outTlb = "$env:TEMP\DeckLinkAPI.tlb"
    Copy-Item $dll $outTlb -Force
    Write-Host "Copied to $outTlb"
    
    # Check if tlbexp or midl available
    $tools = @(
        "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\tlbimp.exe",
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.19041.0\x64\midl.exe"
    )
    foreach ($t in $tools) { if (Test-Path $t) { Write-Host "Found: $t" } }
}
