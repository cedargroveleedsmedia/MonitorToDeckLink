// IDeckLinkOutput - GUID 1A8077F1-9FE2-4533-8147-2294305E253F
// Method order from TLB reflection on installed DeckLinkAPI64.dll (Desktop Video 15.1)
// vtable slot = method_index + 3 (for IUnknown QI/AddRef/Release)
//
//  method[0]  DoesSupportVideoMode          = slot 3
//  method[1]  GetDisplayMode                = slot 4
//  method[2]  GetDisplayModeIterator        = slot 5
//  method[3]  SetScreenPreviewCallback      = slot 6
//  method[4]  EnableVideoOutput             = slot 7  ✓ S_OK confirmed
//  method[5]  DisableVideoOutput            = slot 8
//  method[6]  CreateVideoFrame              = slot 9
//  method[7]  CreateVideoFrameWithBuffer    = slot 10
//  method[8]  RowBytesForPixelFormat        = slot 11
//  method[9]  CreateAncillaryData           = slot 12
//  method[10] DisplayVideoFrameSync         = slot 13
//  method[11] ScheduleVideoFrame            = slot 14
//  method[12] SetScheduledFrameCompletionCallback = slot 15
//  method[13] GetBufferedVideoFrameCount    = slot 16
//  method[14] EnableAudioOutput             = slot 17
//  method[15] DisableAudioOutput            = slot 18
//  method[16] WriteAudioSamplesSync         = slot 19
//  method[17] BeginAudioPreroll             = slot 20
//  method[18] EndAudioPreroll               = slot 21
//  method[19] ScheduleAudioSamples          = slot 22
//  method[20] GetBufferedAudioSampleFrameCount = slot 23
//  method[21] FlushBufferedAudioSamples     = slot 24
//  method[22] SetAudioCallback              = slot 25
//  method[23] StartScheduledPlayback        = slot 26
//  method[24] StopScheduledPlayback         = slot 27

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

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale)
        {
            // ScheduleVideoFrame needs IDeckLinkVideoFrame, QI from IDeckLinkMutableVideoFrame
            Guid videoFrameGuid = new Guid("3F716FE0-F023-4111-BE5D-EF4414C05B17");
            int qhr = Marshal.QueryInterface(frame, ref videoFrameGuid, out IntPtr vf);
            if (qhr != 0) return qhr; // QI failed, return error
            int hr = Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[14])(_ptr, vf, time, dur, scale);
            Marshal.Release(vf);
            return hr;
        }

        public int SetFrameCallback(IntPtr cb) =>
            Marshal.GetDelegateForFunctionPointer<SetCallbackDel>((IntPtr)_vt[15])(_ptr, cb);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[16])(_ptr, out count);

        public int EnableAudioOutput() =>
            Marshal.GetDelegateForFunctionPointer<EnableAudioOutputDel>((IntPtr)_vt[17])(_ptr, 48000, 16, 2, 1);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[26])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[27])(_ptr, stop, out _, scale);
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
