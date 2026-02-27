# Find OBS decklink plugin and check what it does
$obsPlugin = "C:\Program Files\obs-studio\obs-plugins\64bit\decklink-output-ui.dll"
$obsCore   = "C:\Program Files\obs-studio\obs-plugins\64bit\decklink.dll"

foreach ($p in @($obsPlugin, $obsCore)) {
    if (Test-Path $p) { Write-Host "Found: $p" }
    else { Write-Host "Not found: $p" }
}

# More important: load the DeckLink COM object and probe vtable
# using the SAME method OBS uses (BE2D9020 = v14_2_1 per TLB reflection)
# Then call DoesSupportVideoMode at slot 3 to verify slot mapping

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class VtableProbe {
    [DllImport("ole32.dll")] public static extern int CoCreateInstance(ref Guid clsid, IntPtr inner, uint ctx, ref Guid iid, out IntPtr ppv);
    
    public static unsafe void DumpAndProbe() {
        // CoCreateInstance for DeckLink iterator
        Guid clsid = new Guid("9593C2EB-5A35-4155-AFAD-85B46FD94E40"); // CDeckLinkIterator
        Guid iid   = new Guid("7BECFA7C-056B-4F00-86D5-6916A27BD9F1"); // IDeckLinkIterator
        IntPtr iter;
        int hr = CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out iter);
        Console.WriteLine($"CoCreateInstance iterator hr=0x{hr:X8} ptr=0x{iter:X}");
        if (hr != 0) return;
        
        void** ivt = *(void***)iter;
        // IDeckLinkIterator::Next is method [0] = vtable slot 3
        // delegate: int Next(IntPtr self, out IntPtr device)
        var next = Marshal.GetDelegateForFunctionPointer<NextDel>((IntPtr)ivt[3]);
        
        IntPtr dev;
        hr = next(iter, out dev);
        Console.WriteLine($"Iterator.Next hr=0x{hr:X8} dev=0x{dev:X}");
        if (hr != 0 || dev == IntPtr.Zero) return;
        
        // Now QI for IDeckLinkOutput with CURRENT guid (1A8077F1)
        Guid outGuid = new Guid("1A8077F1-9FE2-4533-8147-2294305E253F");
        IntPtr outPtr;
        hr = Marshal.QueryInterface(dev, ref outGuid, out outPtr);
        Console.WriteLine($"QI 1A8077F1 hr=0x{hr:X8}");
        
        // Also try BE2D9020
        Guid outGuid2 = new Guid("BE2D9020-461E-442F-84B7-E949CB953B9D");
        IntPtr outPtr2;
        hr = Marshal.QueryInterface(dev, ref outGuid2, out outPtr2);
        Console.WriteLine($"QI BE2D9020 hr=0x{hr:X8}");
        
        // Use whichever succeeded - dump vtable
        IntPtr use = outPtr != IntPtr.Zero ? outPtr : outPtr2;
        if (use == IntPtr.Zero) { Console.WriteLine("No output interface"); return; }
        
        void** vt = *(void***)use;
        Console.WriteLine("Vtable:");
        for (int i = 0; i <= 27; i++)
            Console.WriteLine($"  [{i:D2}] = 0x{(IntPtr)vt[i]:X}");
        
        // Probe slot 3 (DoesSupportVideoMode) - safe read-only call
        // signature: (IntPtr self, int mode, int w, int h, int rateMode, int flags, out int support, out IntPtr dispMode)
        // Actually just probe GetDisplayModeIterator at slot 5 - returns an iterator object
        Console.WriteLine("Probing slot 5 (GetDisplayModeIterator)...");
        var getModeIter = Marshal.GetDelegateForFunctionPointer<GetModeIterDel>((IntPtr)vt[5]);
        IntPtr modeIter;
        hr = getModeIter(use, 1, out modeIter); // direction=output
        Console.WriteLine($"  slot5 hr=0x{hr:X8} ptr=0x{modeIter:X}");
        
        Marshal.Release(use);
        Marshal.Release(dev);
        Marshal.Release(iter);
    }
    
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] public delegate int NextDel(IntPtr self, out IntPtr dev);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] public delegate int GetModeIterDel(IntPtr self, int dir, out IntPtr iter);
}
"@
[VtableProbe]::DumpAndProbe()
