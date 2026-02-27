# Use TypeLib COM API directly to enumerate IDeckLinkOutput methods
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00020402-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ITypeLib2 {
    int GetTypeInfoCount();
    void GetTypeInfo(int index, out IntPtr ppTI);
    void GetTypeInfoType(int index, out int pTKind);
    void GetTypeInfoOfGuid(ref Guid guid, out IntPtr ppTInfo);
    void GetLibAttr(out IntPtr ppTLibAttr);
    void GetTypeComp(out IntPtr ppTComp);
    void GetDocumentation(int index, [MarshalAs(UnmanagedType.BStr)] out string pbstrName, [MarshalAs(UnmanagedType.BStr)] out string pbstrDocString, out int pdwHelpContext, [MarshalAs(UnmanagedType.BStr)] out string pbstrHelpFile);
    void IsName([MarshalAs(UnmanagedType.LPWStr)] string szNameBuf, int lHashVal, out bool pfName);
    void FindName([MarshalAs(UnmanagedType.LPWStr)] string szNameBuf, int lHashVal, IntPtr[] ppTInfo, int[] rgMemId, ref short pcFound);
    void ReleaseTLibAttr(IntPtr pTLibAttr);
}

[ComImport, Guid("00020403-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ITypeInfo {
    void GetTypeAttr(out IntPtr ppTypeAttr);
    void GetTypeComp(out IntPtr ppTComp);
    void GetFuncDesc(int index, out IntPtr ppFuncDesc);
    void GetVarDesc(int index, out IntPtr ppVarDesc);
    void GetNames(int memid, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] string[] rgBstrNames, int cMaxNames, out int pcNames);
    void GetRefTypeOfImplType(int index, out int href);
    void GetImplTypeFlags(int index, out int pImplTypeFlags);
    void GetIDsOfNames(string[] rgszNames, int cNames, int[] pMemId);
    void Invoke(object pvInstance, int memid, short wFlags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, out object pVarResult, out System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, out int puArgErr);
    void GetDocumentation(int memid, [MarshalAs(UnmanagedType.BStr)] out string pbstrName, [MarshalAs(UnmanagedType.BStr)] out string pbstrDocString, out int pdwHelpContext, [MarshalAs(UnmanagedType.BStr)] out string pbstrHelpFile);
    void GetDllEntry(int memid, System.Runtime.InteropServices.ComTypes.INVOKEKIND invKind, [MarshalAs(UnmanagedType.BStr)] out string pbstrDllName, [MarshalAs(UnmanagedType.BStr)] out string pbstrName, out short pwOrdinal);
    void GetRefTypeInfo(int hRefType, out IntPtr ppTInfo);
    void AddressOfMember(int memid, System.Runtime.InteropServices.ComTypes.INVOKEKIND invKind, out IntPtr ppv);
    void CreateInstance(object pUnkOuter, ref Guid riid, out object ppvObj);
    void GetMops(int memid, [MarshalAs(UnmanagedType.BStr)] out string pbstrMops);
    void GetContainingTypeLib(out IntPtr ppTLib, out int pIndex);
    void ReleaseTypeAttr(IntPtr pTypeAttr);
    void ReleaseFuncDesc(IntPtr pFuncDesc);
    void ReleaseVarDesc(IntPtr pVarDesc);
}

public class TlbNative {
    [DllImport("oleaut32.dll", CharSet=CharSet.Unicode)]
    public static extern int LoadTypeLibEx(string file, int regKind, out IntPtr ppTLib);
}
"@

$dll = "C:\Program Files\Blackmagic Design\Blackmagic Desktop Video\DeckLinkAPI64.dll"
$pTLib = [IntPtr]::Zero
$hr = [TlbNative]::LoadTypeLibEx($dll, 1, [ref]$pTLib)
Write-Host "LoadTypeLibEx hr=0x$($hr.ToString('X8'))"
if ($hr -ne 0) { exit }

$typeLib = [System.Runtime.InteropServices.Marshal]::GetObjectForIUnknown($pTLib) -as [ITypeLib2]
$count = $typeLib.GetTypeInfoCount()
Write-Host "TypeLib has $count types"

for ($i = 0; $i -lt $count; $i++) {
    $name = ""; $doc = ""; $ctx = 0; $help = ""
    $typeLib.GetDocumentation($i, [ref]$name, [ref]$doc, [ref]$ctx, [ref]$help)
    
    if ($name -like "*IDeckLinkOutput*") {
        Write-Host "`n=== $name (index $i) ==="
        $pTI = [IntPtr]::Zero
        $typeLib.GetTypeInfo($i, [ref]$pTI)
        $ti = [System.Runtime.InteropServices.Marshal]::GetObjectForIUnknown($pTI) -as [ITypeInfo]
        
        $pTA = [IntPtr]::Zero
        $ti.GetTypeAttr([ref]$pTA)
        
        # Read cFuncs from TYPEATTR (offset 16 = cFuncs as short)
        $cFuncs = [System.Runtime.InteropServices.Marshal]::ReadInt16($pTA, 16)
        Write-Host "  cFuncs=$cFuncs"
        
        for ($f = 0; $f -lt $cFuncs; $f++) {
            $pFD = [IntPtr]::Zero
            $ti.GetFuncDesc($f, [ref]$pFD)
            # FUNCDESC: memid(4) + ... + oVft(offset 12, short) + ...
            $oVft = [System.Runtime.InteropServices.Marshal]::ReadInt16($pFD, 12)
            $memid = [System.Runtime.InteropServices.Marshal]::ReadInt32($pFD, 0)
            $fname = ""; $fdoc = ""; $fctx = 0; $fhelp = ""
            $ti.GetDocumentation($memid, [ref]$fname, [ref]$fdoc, [ref]$fctx, [ref]$fhelp)
            $slot = $oVft / 8  # 64-bit: pointer size = 8
            Write-Host "  slot[$slot] oVft=$oVft  $fname"
            $ti.ReleaseFuncDesc($pFD)
        }
        $ti.ReleaseTypeAttr($pTA)
    }
}
Write-Host "`nDone."
