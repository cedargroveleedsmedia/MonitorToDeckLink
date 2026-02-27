// Pure vtable P/Invoke for IDeckLinkOutput_v14_2_1 (GUID: 1A8077F1-9FE2-4533-8147-2294305E253F)
// Method order confirmed from tlbimp reflection on Desktop Video 15.1 DeckLinkAPI64.dll
//
// IUnknown:                [0]=QI [1]=AddRef [2]=Release
// IDeckLinkOutput_v14_2_1 interface methods (add 3 for vtable slot):
//  [0] DoesSupportVideoMode      = slot 3
//  [1] GetDisplayMode            = slot 4
//  [2] GetDisplayModeIterator    = slot 5
//  [3] SetScreenPreviewCallback  = slot 6
//  [4] EnableVideoOutput         = slot 7  *** CONFIRMED ***
//  [5] DisableVideoOutput        = slot 8
//  [6] SetVideoOutputFrameMemoryAllocator = slot 9  (was incorrectly used as CreateVideoFrame)
//  [7] CreateVideoFrame          = slot 10 *** FIXED ***
//  [8] CreateAncillaryData       = slot 11
//  [9] DisplayVideoFrameSync     = slot 12 *** FIXED ***
// [10] ScheduleVideoFrame        = slot 13 *** FIXED ***
// [11] SetScheduledFrameCompletionCallback = slot 14
// [12] GetBufferedVideoFrameCount = slot 15
// [13] EnableAudioOutput         = slot 16
// [14] DisableAudioOutput        = slot 17
// [15] WriteAudioSamplesSync     = slot 18
// [16] BeginAudioPreroll         = slot 19
// [17] EndAudioPreroll           = slot 20
// [18] ScheduleAudioSamples      = slot 21
// [19] GetBufferedAudioSampleFrameCount = slot 22
// [20] FlushBufferedAudioSamples = slot 23
// [21] SetAudioCallback          = slot 24
// [22] StartScheduledPlayback    = slot 25 *** FIXED ***
// [23] StopScheduledPlayback     = slot 26 *** FIXED ***
// [24] IsScheduledPlaybackRunning= slot 27
// [25] GetScheduledStreamTime    = slot 28
// ...

using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    internal sealed unsafe class DeckLinkOutput : IDisposable
    {
        private IntPtr _ptr;
        private void** _vt;

        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint ReleaseDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint AddRefDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int EnableVideoOutputDel(IntPtr self, int mode, int flags);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int DisableVideoOutputDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int CreateVideoFrameDel(IntPtr self, int w, int h, int rb, int fmt, int flags, out IntPtr frame);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int DisplayVideoFrameSyncDel(IntPtr self, IntPtr frame);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int ScheduleVideoFrameDel(IntPtr self, IntPtr frame, long time, long dur, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBufferedCountDel(IntPtr self, out uint count);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StartPlaybackDel(IntPtr self, long startTime, long scale, double speed);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StopPlaybackDel(IntPtr self, long stopTime, out long actualStop, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBytesDel(IntPtr self, out IntPtr buffer);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StartAccessDel(IntPtr self, int accessType);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int EndAccessDel(IntPtr self, int accessType);

        const int bmdBufferAccessWrite = 2;

        public DeckLinkOutput(IntPtr ptr)
        {
            _ptr = ptr;
            _vt  = *(void***)ptr;
            Marshal.GetDelegateForFunctionPointer<AddRefDel>((IntPtr)_vt[1])(_ptr);
        }

        public int EnableVideoOutput(int mode, int flags) =>
            Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vt[7])(_ptr, mode, flags);

        public int DisableVideoOutput() =>
            Marshal.GetDelegateForFunctionPointer<DisableVideoOutputDel>((IntPtr)_vt[8])(_ptr);

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[10])(_ptr, w, h, rb, fmt, flags, out frame);

        public int DisplayVideoFrameSync(IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<DisplayVideoFrameSyncDel>((IntPtr)_vt[12])(_ptr, frame);

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[13])(_ptr, frame, time, dur, scale);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[15])(_ptr, out count);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[25])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[26])(_ptr, stop, out _, scale);
            return 0;
        }

        public void GetFrameBytes(IntPtr frame, out IntPtr bytes, System.Action<string> log)
        {
            bytes = IntPtr.Zero;
            Guid bufGuid = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qiHr = Marshal.QueryInterface(frame, ref bufGuid, out IntPtr bufPtr);
            if (qiHr != 0 || bufPtr == IntPtr.Zero) { log($"QI IDeckLinkVideoBuffer failed: 0x{qiHr:X8}"); return; }
            void** bvt = *(void***)bufPtr;
            Marshal.GetDelegateForFunctionPointer<StartAccessDel>((IntPtr)bvt[4])(bufPtr, bmdBufferAccessWrite);
            Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[3])(bufPtr, out bytes);
            _lastBufPtr = bufPtr;
            _lastBufVt  = bvt;
        }

        private IntPtr _lastBufPtr;
        private void** _lastBufVt;

        public void EndFrameAccess()
        {
            if (_lastBufPtr != IntPtr.Zero)
            {
                Marshal.GetDelegateForFunctionPointer<EndAccessDel>((IntPtr)_lastBufVt[5])(_lastBufPtr, bmdBufferAccessWrite);
                Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_lastBufVt[2])(_lastBufPtr);
                _lastBufPtr = IntPtr.Zero;
            }
        }

        public void ReleaseFrame(IntPtr frame)
        {
            void** fvt = *(void***)frame;
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)fvt[2])(frame);
        }

        public string DumpVtable()
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < 28; i++)
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
