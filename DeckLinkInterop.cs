using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    [ComImport, Guid("50FB36CD-3063-4B73-BDBB-958087F2D8BA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkIterator2
    {
        [PreserveSig] int Next([MarshalAs(UnmanagedType.Interface)] out IDeckLink2 deckLink);
    }

    [ComImport, Guid("C418FBDD-0587-48ED-8FE5-640F0A14AF91"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLink2
    {
        [PreserveSig] int GetModelName([MarshalAs(UnmanagedType.BStr)] out string name);
        [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.BStr)] out string name);
    }

    [ComImport, Guid("1A8077F1-9FE2-4533-8147-2294305E253F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkOutput2
    {
        [PreserveSig] int DoesSupportVideoMode(int displayMode, int pixelFormat, int flags, out int supported, [MarshalAs(UnmanagedType.Interface)] out object displayModeObj);
        [PreserveSig] int GetDisplayMode(int displayMode, [MarshalAs(UnmanagedType.Interface)] out object displayModeObj);
        [PreserveSig] int GetDisplayModeIterator([MarshalAs(UnmanagedType.Interface)] out object iterator);
        [PreserveSig] int SetScreenPreviewCallback([MarshalAs(UnmanagedType.Interface)] object callback);
        [PreserveSig] int EnableVideoOutput(int displayMode, int flags);
        [PreserveSig] int DisableVideoOutput();
        [PreserveSig] int CreateVideoFrame(int width, int height, int rowBytes, int pixelFormat, int flags, [MarshalAs(UnmanagedType.Interface)] out IDeckLinkMutableVideoFrame2 frame);
        [PreserveSig] int ScheduleVideoFrame([MarshalAs(UnmanagedType.Interface)] IDeckLinkMutableVideoFrame2 frame, long displayTime, long displayDuration, long timeScale);
    }

    [ComImport, Guid("CF9EB134-0374-4C5B-95FA-1EC14819FF62"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkMutableVideoFrame2
    {
        int GetWidth();
        int GetHeight();
        int GetRowBytes();
        int GetPixelFormat();
        int GetFlags();
        [PreserveSig] int GetBytes(out IntPtr buffer);
    }

    [ComImport, Guid("BA6C6F44-6DA5-4DCE-94AA-EE2D1372A676")]
    public class CDeckLinkIterator2 { }
}
