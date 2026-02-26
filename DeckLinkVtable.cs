// Pure vtable P/Invoke - bypasses RCW entirely
// Vtable layout from OBS Studio DeckLink SDK (modern Desktop Video 12+):
//   IUnknown:  [0]=QI  [1]=AddRef  [2]=Release
//   IDeckLinkOutput (modern):
//   [3]  DoesSupportVideoMode  (6 params - connection, mode, pixelFmt, convMode, flags, *actualMode, *supported)
//   [4]  GetDisplayMode
//   [5]  GetDisplayModeIterator
//   [6]  SetScreenPreviewCallback
//   [7]  EnableVideoOutput          <-- slot 7 in modern SDK
//   [8]  DisableVideoOutput
//   [9]  SetVideoOutputFrameMemoryAllocator
//   [10] CreateVideoFrame
//   [11] SetAncillaryData
//   [12] SetVideoOutputConversionMode
//   [13] ScheduleVideoFrame
//   [14] GetBufferedVideoFrameCount
//   [15] StartScheduledPlayback
//   [16] StopScheduledPlayback
//   [17] IsScheduledPlaybackRunning
//   [18] GetScheduledStreamTime
//   [19] GetReferenceStatus
//   [20] EnableAudioOutput
//   [21] DisableAudioOutput
//   ...

using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
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
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBytesDel(IntPtr self, out IntPtr buffer);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetWidthDel(IntPtr self);
        // Frame vtable dump helper
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint AddRefDel2(IntPtr self);

        public DeckLinkOutput(IntPtr ptr)
        {
            _ptr = ptr;
            _vt  = *(void***)ptr;
            var ar = Marshal.GetDelegateForFunctionPointer<AddRefDel>((IntPtr)_vt[1]);
            ar(_ptr);
        }

        // Modern SDK slot 7
        public int EnableVideoOutput(int mode, int flags) =>
            Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vt[7])(_ptr, mode, flags);

        // Modern SDK slot 8
        public int DisableVideoOutput() =>
            Marshal.GetDelegateForFunctionPointer<DisableVideoOutputDel>((IntPtr)_vt[8])(_ptr);

        // Try slot 9 - vtable shows [10] < [09] suggesting different ordering
        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        public string DumpFrameVtable(IntPtr frame)
        {
            void** fvt = *(void***)frame;
            var sb = new System.Text.StringBuilder("Frame vtable:\n");
            for (int i = 0; i < 16; i++)
                sb.AppendLine($"  [{i:D2}] = 0x{(IntPtr)fvt[i]:X}");
            return sb.ToString();
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StartAccessDel(IntPtr self, int accessType);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int EndAccessDel(IntPtr self, int accessType);

        // bmdBufferAccessRead=1, bmdBufferAccessWrite=2
        const int bmdBufferAccessWrite = 2;

        public void GetFrameBytes(IntPtr frame, out IntPtr bytes, System.Action<string> log)
        {
            bytes = IntPtr.Zero;
            Guid bufGuid = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qiHr = Marshal.QueryInterface(frame, ref bufGuid, out IntPtr bufPtr);
            if (qiHr != 0 || bufPtr == IntPtr.Zero)
            {
                log($"QI IDeckLinkVideoBuffer failed: 0x{qiHr:X8}");
                return;
            }
            void** bvt = *(void***)bufPtr;
            // IDeckLinkVideoBuffer: [0]QI [1]AddRef [2]Release [3]GetBytes [4]StartAccess [5]EndAccess
            // Must call StartAccess before GetBytes
            int saHr = Marshal.GetDelegateForFunctionPointer<StartAccessDel>((IntPtr)bvt[4])(bufPtr, bmdBufferAccessWrite);
            log($"StartAccess hr=0x{saHr:X8}");
            if (saHr == 0)
            {
                int hr = Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[3])(bufPtr, out bytes);
                log($"GetBytes hr=0x{hr:X8} bytes=0x{bytes:X}");
            }
            // EndAccess is called after we are done writing (in WriteFrameBytes)
            // Store bufPtr for EndAccess - caller must call EndFrameAccess
            _lastBufPtr = bufPtr;
            _lastBufVt  = bvt;
        }

        private IntPtr _lastBufPtr;
        private void** _lastBufVt;

        public void EndFrameAccess(System.Action<string> log)
        {
            if (_lastBufPtr != IntPtr.Zero)
            {
                int hr = Marshal.GetDelegateForFunctionPointer<EndAccessDel>((IntPtr)_lastBufVt[5])(_lastBufPtr, bmdBufferAccessWrite);
                log($"EndAccess hr=0x{hr:X8}");
                Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_lastBufVt[2])(_lastBufPtr);
                _lastBufPtr = IntPtr.Zero;
            }
        }

        public void ReleaseFrame(IntPtr frame)
        {
            void** fvt = *(void***)frame;
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)fvt[2])(frame);
        }

        // Modern SDK slot 13
        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[13])(_ptr, frame, time, dur, scale);

        // Modern SDK slot 14
        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[14])(_ptr, out count);

        // Modern SDK slot 15
        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[15])(_ptr, start, scale, speed);

        // Modern SDK slot 16
        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[16])(_ptr, stop, out _, scale);
            return 0;
        }

        public string DumpVtable()
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < 22; i++)
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
