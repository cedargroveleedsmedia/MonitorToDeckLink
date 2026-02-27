// IDeckLinkOutput (current, GUID CC5C8A6E-3F2F-4B3A-87EA-FD78AF300564)
// TLB method order (add 3 for IUnknown to get vtable slot):
//  [0] DoesSupportVideoMode      = slot 3
//  [1] GetDisplayMode            = slot 4
//  [2] GetDisplayModeIterator    = slot 5
//  [3] SetScreenPreviewCallback  = slot 6
//  [4] EnableVideoOutput         = slot 7
//  [5] DisableVideoOutput        = slot 8
//  [6] CreateVideoFrame          = slot 9
//  [7] CreateVideoFrameWithBuffer= slot 10
//  [8] RowBytesForPixelFormat    = slot 11
//  [9] CreateAncillaryData       = slot 12
// [10] DisplayVideoFrameSync     = slot 13
// [11] ScheduleVideoFrame        = slot 14
// [12] SetScheduledFrameCompletionCallback = slot 15
// [13] GetBufferedVideoFrameCount= slot 16
// ...
// [22] StartScheduledPlayback    = slot 25
// [23] StopScheduledPlayback     = slot 26
//
// IDeckLinkMutableVideoFrame (current GUID) method order:
//  [0] GetWidth  [1] GetHeight  [2] GetRowBytes  [3] GetPixelFormat
//  [4] GetFlags  [5] GetTimecode  [6] GetAncillaryData
//  [7] SetFlags  [8] SetTimecode  [9] SetTimecodeFromComponents
// [10] SetAncillaryData  [11] SetTimecodeUserBits  [12] SetInterfaceProvider
// GetBytes is on IDeckLinkVideoBuffer (separate QI, GUID CCB4B64A):
//  [0] GetBytes  [1] StartAccess  [2] EndAccess  => vtable slots 3,4,5

using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    internal sealed unsafe class DeckLinkOutput : IDisposable
    {
        private IntPtr _ptr;
        private void** _vt;

        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint   ReleaseDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint   AddRefDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    EnableVideoOutputDel(IntPtr self, int mode, int flags);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    DisableVideoOutputDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    CreateVideoFrameDel(IntPtr self, int w, int h, int rb, int fmt, int flags, out IntPtr frame);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    ScheduleVideoFrameDel(IntPtr self, IntPtr frame, long time, long dur, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    GetBufferedCountDel(IntPtr self, out uint count);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    StartPlaybackDel(IntPtr self, long startTime, long scale, double speed);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    StopPlaybackDel(IntPtr self, long stopTime, out long actualStop, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    EnableAudioOutputDel(IntPtr self, int sampleRate, int sampleType, uint channelCount, int streamType);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    GetBytesDel(IntPtr self, out IntPtr buffer);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    SetCallbackDel(IntPtr self, IntPtr callback);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    StartAccessDel(IntPtr self, int accessType);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int    EndAccessDel(IntPtr self, int accessType);

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

        // EnableAudioOutput at slot 16 (interface method 13)
        // bmdAudioSampleRate48kHz=48000, bmdAudioSampleType16bitInteger=16, bmdAudioOutputStreamTimestamped=1
        public int EnableAudioOutput() =>
            Marshal.GetDelegateForFunctionPointer<EnableAudioOutputDel>((IntPtr)_vt[16])(_ptr, 48000, 16, 2, 1);

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[14])(_ptr, frame, time, dur, scale);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[16])(_ptr, out count);

        // SetScheduledFrameCompletionCallback at slot 15 (interface method 12)
        public int SetFrameCallback(IntPtr callbackPtr) =>
            Marshal.GetDelegateForFunctionPointer<SetCallbackDel>((IntPtr)_vt[15])(_ptr, callbackPtr);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[25])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[26])(_ptr, stop, out _, scale);
            return 0;
        }

        // IDeckLinkVideoBuffer (QI with GUID CCB4B64A): GetBytes[3] StartAccess[4] EndAccess[5]
        public int GetFrameBytes(IntPtr frame, out IntPtr bytes, System.Action<string> log)
        {
            bytes = IntPtr.Zero;
            Guid g = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qhr = Marshal.QueryInterface(frame, ref g, out IntPtr buf);
            log($"QI IDeckLinkVideoBuffer: 0x{qhr:X8}");
            if (qhr != 0) return qhr;
            void** bvt = *(void***)buf;
            int saHr = Marshal.GetDelegateForFunctionPointer<StartAccessDel>((IntPtr)bvt[4])(buf, 2);
            log($"StartAccess: 0x{saHr:X8}");
            int gbHr = Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[3])(buf, out bytes);
            log($"GetBytes: 0x{gbHr:X8} -> 0x{bytes:X}");
            _bufPtr = buf; _bufVt = bvt;
            return gbHr;
        }

        private IntPtr _bufPtr; private void** _bufVt;

        public void EndFrameAccess()
        {
            if (_bufPtr == IntPtr.Zero) return;
            Marshal.GetDelegateForFunctionPointer<EndAccessDel>((IntPtr)_bufVt[5])(_bufPtr, 2);
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_bufVt[2])(_bufPtr);
            _bufPtr = IntPtr.Zero;
        }

        public void ReleaseFrame(IntPtr frame)
        {
            void** fvt = *(void***)frame;
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)fvt[2])(frame);
        }

        public string DumpVtable()
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i <= 27; i++)
                sb.AppendLine($"  [{i:D2}] = 0x{(IntPtr)_vt[i]:X}");
            return sb.ToString();
        }

        public void Dispose()
        {
            if (_ptr == IntPtr.Zero) return;
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_vt[2])(_ptr);
            _ptr = IntPtr.Zero;
        }
    }
}
