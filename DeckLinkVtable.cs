// Pure vtable P/Invoke - bypasses RCW entirely
// All DeckLink COM calls go through raw function pointers

using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    /// <summary>
    /// Wraps an IDeckLinkOutput COM pointer and calls methods via raw vtable.
    /// Vtable layout from DeckLink SDK 12.x Win64 DeckLinkAPI.h:
    ///   IUnknown:  [0]=QI  [1]=AddRef  [2]=Release
    ///   IDeckLinkOutput:
    ///   [3]  DoesSupportVideoMode
    ///   [4]  GetDisplayModeIterator
    ///   [5]  SetScreenPreviewCallback
    ///   [6]  EnableVideoOutput
    ///   [7]  DisableVideoOutput
    ///   [8]  SetVideoOutputFrameMemoryAllocator
    ///   [9]  CreateVideoFrame
    ///   [10] SetAncillaryData
    ///   [11] SetVideoOutputConversionMode
    ///   [12] ScheduleVideoFrame
    ///   [13] GetBufferedVideoFrameCount
    ///   [14] StartScheduledPlayback
    ///   [15] StopScheduledPlayback
    ///   [16] IsScheduledPlaybackRunning
    ///   [17] GetScheduledStreamTime
    ///   [18] GetReferenceStatus
    ///   [19] EnableAudioOutput
    ///   [20] DisableAudioOutput
    ///   [21] WriteAudioSamplesSync
    ///   [22] BeginAudioPreroll
    ///   [23] EndAudioPreroll
    ///   [24] ScheduleAudioSamples
    ///   [25] GetBufferedAudioSampleFrameCount
    ///   [26] FlushBufferedAudioSamples
    ///   [27] SetAudioCallback
    ///   [28] GetOutputVideoFrameState
    /// </summary>
    internal sealed unsafe class DeckLinkOutput : IDisposable
    {
        private IntPtr _ptr;
        private void** _vt;

        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint AddRefDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint ReleaseDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int EnableVideoOutputDel(IntPtr self, int mode, int flags);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int DisableVideoOutputDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int CreateVideoFrameDel(IntPtr self, int w, int h, int rb, int fmt, int flags, out IntPtr frame);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int ScheduleVideoFrameDel(IntPtr self, IntPtr frame, long time, long dur, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBufferedCountDel(IntPtr self, out uint count);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StartPlaybackDel(IntPtr self, long startTime, long scale, double speed);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StopPlaybackDel(IntPtr self, long stopTime, out long actualStop, long scale);
        // IDeckLinkVideoFrame::GetBytes (on MutableVideoFrame vtable slot 5 = IUnknown[3] + frame[2])
        // IDeckLinkVideoFrame: [0]QI [1]AddRef [2]Release [3]GetWidth [4]GetHeight [5]GetRowBytes [6]GetPixelFormat [7]GetFlags [8]GetBytes
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBytesDel(IntPtr self, out IntPtr buffer);

        public DeckLinkOutput(IntPtr ptr)
        {
            _ptr = ptr;
            _vt  = *(void***)ptr;
            // AddRef
            var ar = Marshal.GetDelegateForFunctionPointer<AddRefDel>((IntPtr)_vt[1]);
            ar(_ptr);
        }

        public int EnableVideoOutput(int mode, int flags) =>
            Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vt[6])(_ptr, mode, flags);

        public int DisableVideoOutput() =>
            Marshal.GetDelegateForFunctionPointer<DisableVideoOutputDel>((IntPtr)_vt[7])(_ptr);

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        public void GetFrameBytes(IntPtr frame, out IntPtr bytes)
        {
            void** fvt = *(void***)frame;
            Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)fvt[8])(frame, out bytes);
        }

        public void ReleaseFrame(IntPtr frame)
        {
            void** fvt = *(void***)frame;
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)fvt[2])(frame);
        }

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[12])(_ptr, frame, time, dur, scale);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[13])(_ptr, out count);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[14])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[15])(_ptr, stop, out _, scale);
            return 0;
        }

        public string DumpVtable()
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < 20; i++)
                sb.AppendLine($"  [{i:D2}] = 0x{(IntPtr)_vt[i]:X}");
            return sb.ToString();
        }

        public void Dispose()
        {
            if (_ptr != IntPtr.Zero)
            {
                Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_vt[2])(_ptr);
                _ptr = IntPtr.Zero;
            }
        }
    }
}
