// IDeckLinkOutput CURRENT interface - GUID 1A8077F1 (Desktop Video 15.x)
// Method order from TLB reflection on DeckLinkInterop.dll (generated from installed DeckLinkAPI64.dll)
// Interface methods (add 3 for IUnknown slots to get vtable slot):
//  [0] DoesSupportVideoMode           = slot 3
//  [1] GetDisplayMode                 = slot 4
//  [2] GetDisplayModeIterator         = slot 5
//  [3] SetScreenPreviewCallback       = slot 6
//  [4] EnableVideoOutput              = slot 7  ✓ confirmed S_OK
//  [5] DisableVideoOutput             = slot 8
//  [6] SetVideoOutputFrameMemoryAllocator = slot 9
//  [7] CreateVideoFrame               = slot 10
//  [8] CreateAncillaryData            = slot 11
//  [9] DisplayVideoFrameSync          = slot 12
// [10] ScheduleVideoFrame             = slot 13
// [11] SetScheduledFrameCompletionCallback = slot 14  ✓ confirmed S_OK
// [12] GetBufferedVideoFrameCount     = slot 15
// [13] StartScheduledPlayback         = slot 16? — needs verification
// ...
// EnableAudioOutput confirmed at slot 16 via S_OK — so audio methods start at 16
// That means StartScheduledPlayback must be after audio methods
// From TLB: audio methods occupy slots 16-22, StartScheduledPlayback = slot 23? 
// Use vtable dump: slots 12+ are near-sequential stubs except real ones
// EnableAudioOutput=16 confirmed. Count audio methods from IDL:
//   EnableAudioOutput, DisableAudioOutput, WriteAudioSamplesSync, 
//   BeginAudioPreroll, EndAudioPreroll, ScheduleAudioSamples,
//   GetBufferedAudioSampleFrameCount, FlushBufferedAudioSamples,
//   SetAudioCallback = 9 audio methods = slots 16-24
// StartScheduledPlayback = slot 25, StopScheduledPlayback = slot 26

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
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  EnableVideoOutputDel(IntPtr self, int mode, int flags);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  DisableVideoOutputDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  CreateVideoFrameDel(IntPtr self, int w, int h, int rb, int fmt, int flags, out IntPtr frame);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  ScheduleVideoFrameDel(IntPtr self, IntPtr frame, long time, long dur, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  SetCallbackDel(IntPtr self, IntPtr cb);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  GetBufferedCountDel(IntPtr self, out uint count);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  StartPlaybackDel(IntPtr self, long startTime, long scale, double speed);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  StopPlaybackDel(IntPtr self, long stopTime, out long actualStop, long scale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  EnableAudioOutputDel(IntPtr self, int sampleRate, int sampleType, uint channelCount, int streamType);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  GetBytesDel(IntPtr self, out IntPtr buffer);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  StartAccessDel(IntPtr self, int accessType);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int  EndAccessDel(IntPtr self, int accessType);

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

        public int EnableAudioOutput() =>
            Marshal.GetDelegateForFunctionPointer<EnableAudioOutputDel>((IntPtr)_vt[16])(_ptr, 48000, 16, 2, 1);

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[13])(_ptr, frame, time, dur, scale);

        public int SetFrameCallback(IntPtr cb) =>
            Marshal.GetDelegateForFunctionPointer<SetCallbackDel>((IntPtr)_vt[14])(_ptr, cb);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[15])(_ptr, out count);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[25])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[26])(_ptr, stop, out _, scale);
            return 0;
        }

        public int GetFrameBytes(IntPtr frame, out IntPtr bytes, System.Action<string> log)
        {
            bytes = IntPtr.Zero;
            Guid g = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qhr = Marshal.QueryInterface(frame, ref g, out IntPtr buf);
            if (qhr != 0) { log($"QI IDeckLinkVideoBuffer failed: 0x{qhr:X8}"); return qhr; }
            void** bvt = *(void***)buf;
            Marshal.GetDelegateForFunctionPointer<StartAccessDel>((IntPtr)bvt[4])(buf, 2);
            int gbHr = Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[3])(buf, out bytes);
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
