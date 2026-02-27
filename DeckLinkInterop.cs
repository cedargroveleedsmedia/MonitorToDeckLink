// Hand-written COM interop for DeckLink - bypasses tlbimp issues entirely
// GUIDs sourced from DeckLink SDK headers and community wrappers

using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    // ── IDeckLinkIterator ─────────────────────────────────────────────────────
    [ComImport, Guid("50FB36CD-3063-4B73-BDBB-958087F2D8BA"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkIterator2
    {
        [PreserveSig] int Next([MarshalAs(UnmanagedType.Interface)] out IDeckLink2 deckLink);
    }

    // ── IDeckLink ─────────────────────────────────────────────────────────────
    [ComImport, Guid("C418FBDD-0587-48ED-8FE5-640F0A14AF91"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLink2
    {
        [PreserveSig] int GetModelName([MarshalAs(UnmanagedType.BStr)] out string name);
        [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.BStr)] out string name);
    }

    // ── IDeckLinkOutput ───────────────────────────────────────────────────────
    // GUID from DeckLink SDK community wrapper (A3EF0963) - different from tlbimp
    [ComImport, Guid("1A8077F1-9FE2-4533-8147-2294305E253F"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkOutput2
    {
        [PreserveSig] int DoesSupportVideoMode(
            int displayMode, int pixelFormat, int flags,
            out int supported, [MarshalAs(UnmanagedType.Interface)] out object displayModeObj);
        [PreserveSig] int GetDisplayModeIterator(
            [MarshalAs(UnmanagedType.Interface)] out object iterator);
        [PreserveSig] int SetScreenPreviewCallback(
            [MarshalAs(UnmanagedType.Interface)] object callback);
        [PreserveSig] int EnableVideoOutput(int displayMode, int flags);
        [PreserveSig] int DisableVideoOutput();
        [PreserveSig] int SetVideoOutputFrameMemoryAllocator(
            [MarshalAs(UnmanagedType.Interface)] object allocator);
        [PreserveSig] int CreateVideoFrame(
            int width, int height, int rowBytes, int pixelFormat, int flags,
            [MarshalAs(UnmanagedType.Interface)] out IDeckLinkMutableVideoFrame2 frame);
        [PreserveSig] int SetAncillaryData(
            [MarshalAs(UnmanagedType.Interface)] object ancillary);
        [PreserveSig] int SetVideoOutputConversionMode(int conversionMode);
        [PreserveSig] int ScheduleVideoFrame(
            [MarshalAs(UnmanagedType.Interface)] IDeckLinkMutableVideoFrame2 frame,
            long displayTime, long displayDuration, long timeScale);
        [PreserveSig] int GetBufferedVideoFrameCount(out uint bufferedFrameCount);
        [PreserveSig] int StartScheduledPlayback(
            long playbackStartTime, long timeScale, double playbackSpeed);
        [PreserveSig] int StopScheduledPlayback(
            long stopPlaybackAtTime, out long actualStopTime, long timeScale);
        [PreserveSig] int IsScheduledPlaybackRunning(out bool active);
        [PreserveSig] int GetScheduledStreamTime(
            long desiredTimeScale, out long streamTime, out double playbackSpeed);
        [PreserveSig] int GetReferenceStatus(out int referenceStatus);
        [PreserveSig] int EnableAudioOutput(
            int sampleRate, int sampleType, uint channelCount, int streamType);
        [PreserveSig] int DisableAudioOutput();
        [PreserveSig] int WriteAudioSamplesSync(
            IntPtr buffer, uint sampleFrameCount, out uint sampleFramesWritten);
        [PreserveSig] int BeginAudioPreroll();
        [PreserveSig] int EndAudioPreroll();
        [PreserveSig] int ScheduleAudioSamples(
            IntPtr buffer, uint sampleFrameCount, long streamTime,
            long timeScale, out uint sampleFramesWritten);
        [PreserveSig] int GetBufferedAudioSampleFrameCount(out uint bufferedSampleFrameCount);
        [PreserveSig] int FlushBufferedAudioSamples();
        [PreserveSig] int SetAudioCallback(
            [MarshalAs(UnmanagedType.Interface)] object callback);
        [PreserveSig] int GetOutputVideoFrameState(
            [MarshalAs(UnmanagedType.Interface)] IDeckLinkMutableVideoFrame2 frame,
            out int state);
    }

    // ── IDeckLinkMutableVideoFrame ────────────────────────────────────────────
    [ComImport, Guid("CF9EB134-0374-4C5B-95FA-1EC14819FF62"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkMutableVideoFrame2
    {
        // IDeckLinkVideoFrame methods
        [PreserveSig] int GetWidth();
        [PreserveSig] int GetHeight();
        [PreserveSig] int GetRowBytes();
        [PreserveSig] int GetPixelFormat();
        [PreserveSig] int GetFlags();
        [PreserveSig] int GetBytes(out IntPtr buffer);
        [PreserveSig] int GetTimecode(int format,
            [MarshalAs(UnmanagedType.Interface)] out object timecode);
        [PreserveSig] int GetAncillaryData(
            [MarshalAs(UnmanagedType.Interface)] out object ancillary);
        // IDeckLinkMutableVideoFrame methods
        [PreserveSig] int SetFlags(int newFlags);
        [PreserveSig] int SetTimecode(int format,
            [MarshalAs(UnmanagedType.Interface)] object timecode);
        [PreserveSig] int SetTimecodeFromComponents(
            int format, byte hours, byte minutes, byte seconds, byte frames, int flags);
        [PreserveSig] int SetAncillaryData(
            [MarshalAs(UnmanagedType.Interface)] object ancillary);
        [PreserveSig] int SetTimecodeUserBits(int format, uint userBits);
        [PreserveSig] int SetInterfaceProvider(
            [MarshalAs(UnmanagedType.Interface)] object provider);
    }

    // ── CDeckLinkIterator CoClass ──────────────────────────────────────────────
    [ComImport, Guid("BA6C6F44-6DA5-4DCE-94AA-EE2D1372A676")]
    public class CDeckLinkIterator2 { }
}
