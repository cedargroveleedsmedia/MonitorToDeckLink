using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    [ComImport, Guid("50FB36CD-3063-4B73-BDBB-958087F2D8BA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkIterator2
    {
        [PreserveSig] int Next(out IntPtr deckLink);
    }

    [ComImport, Guid("C418FBDD-0587-48ED-8FE5-640F0A14AF91"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLink2
    {
        [PreserveSig] int GetModelName([MarshalAs(UnmanagedType.BStr)] out string name);
    }

    [ComImport, Guid("CF9EB134-0374-4C5B-95FA-1EC14819FF62"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkVideoFrame
    {
        int GetWidth();
        int GetHeight();
        int GetRowBytes();
        int GetPixelFormat();
        int GetFlags();
        [PreserveSig] int GetBytes(out IntPtr buffer);
    }
}
