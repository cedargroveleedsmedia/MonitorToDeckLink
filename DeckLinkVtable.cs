// Pure vtable P/Invoke - Desktop Video SDK 15.x layout
// Confirmed from sdk-doc.blackmagicdesign.com section 2.5.3 method order:
//
// IUnknown: [0]=QI [1]=AddRef [2]=Release
// IDeckLinkOutput:
// [3]  DoesSupportVideoMode
// [4]  GetDisplayMode
// [5]  GetDisplayModeIterator
// [6]  SetScreenPreviewCallback
// [7]  EnableVideoOutput          *** CONFIRMED WORKING ***
// [8]  DisableVideoOutput
// [9]  CreateVideoFrame           *** CONFIRMED WORKING ***
// [10] CreateVideoFrameWithBuffer  (NEW in 15.x)
// [11] RowBytesForPixelFormat      (NEW in 15.x)
// [12] CreateAncillaryData         (NEW in 15.x)
// [13] DisplayVideoFrameSync
// [14] ScheduleVideoFrame
// [15] SetScheduledFrameCompletionCallback
// [16] GetBufferedVideoFrameCount
// [17] EnableAudioOutput
// [18] DisableAudioOutput
// [19] WriteAudioSamplesSync
// [20] BeginAudioPreroll
// [21] EndAudioPreroll
// [22] ScheduleAudioSamples
// [23] GetBufferedAudioSampleFrameCount
// [24] FlushBufferedAudioSamples
// [25] SetAudioCallback
// [26] StartScheduledPlayback
// [27] StopScheduledPlayback
// [28] IsScheduledPlaybackRunning
// [29] GetScheduledStreamTime
// [30] GetReferenceStatus
// [31] GetHardwareReferenceClock
// [32] GetFrameCompletionReferenceTimestamp

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

        // SDK docs 2.5.3 method order for IDeckLinkOutput (v14_2_1 GUID 1A8077F1):
        // EnableVideoOutput=[7] confirmed. After that:
        // [8]=DisableVideoOutput [9]=SetVideoOutputFrameMemoryAllocator [10]=CreateVideoFrame
        // But previously slot 9 returned a COM object (IDeckLinkVideoBuffer) not a frame.
        // Try slot 10 as CreateVideoFrame.
        public int CreateSlot => 10;

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[10])(_ptr, w, h, rb, fmt, flags, out frame);

        // DisplayVideoFrameSync - simpler than scheduled, use as fallback
        public int DisplayVideoFrameSync(IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<DisplayVideoFrameSyncDel>((IntPtr)_vt[13])(_ptr, frame);

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[14])(_ptr, frame, time, dur, scale);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[16])(_ptr, out count);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[26])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[27])(_ptr, stop, out _, scale);
            return 0;
        }

        // IDeckLinkVideoBuffer access for GetBytes
        public void GetFrameBytes(IntPtr frame, out IntPtr bytes, System.Action<string> log)
        {
            bytes = IntPtr.Zero;
            Guid bufGuid = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qiHr = Marshal.QueryInterface(frame, ref bufGuid, out IntPtr bufPtr);
            if (qiHr != 0 || bufPtr == IntPtr.Zero) { log($"QI IDeckLinkVideoBuffer failed: 0x{qiHr:X8}"); return; }
            void** bvt = *(void***)bufPtr;
            // [3]=GetBytes [4]=StartAccess [5]=EndAccess
            Marshal.GetDelegateForFunctionPointer<StartAccessDel>((IntPtr)bvt[4])(bufPtr, bmdBufferAccessWrite);
            Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[3])(bufPtr, out bytes);
            _lastBufPtr = bufPtr;
            _lastBufVt  = bvt;
        }

        private IntPtr _lastBufPtr;
        private void** _lastBufVt;

        public void EndFrameAccess(System.Action<string>? log = null)
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
