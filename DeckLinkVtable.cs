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

        public Action<string>? Logger { get; set; }

        public DeckLinkOutput(IntPtr ptr)
        {
            _ptr = ptr;
            _vt  = *(void***)ptr;
            Marshal.GetDelegateForFunctionPointer<AddRefDel>((IntPtr)_vt[1])(_ptr);
        }

        public int EnableVideoOutput(int mode, int flags) =>
            Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vt[7])(_ptr, mode, flags);

        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[14])(_ptr, frame, time, dur, scale);

        public int SetFrameCallback(IntPtr cb) =>
            Marshal.GetDelegateForFunctionPointer<SetCallbackDel>((IntPtr)_vt[15])(_ptr, cb);

        public int EnableAudioOutput() =>
            Marshal.GetDelegateForFunctionPointer<EnableAudioOutputDel>((IntPtr)_vt[17])(_ptr, 48000, 16, 2, 1);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[26])(_ptr, start, scale, speed);

        public int StopScheduledPlayback(long stop, long scale)
        {
            if (_ptr == IntPtr.Zero) return 0;
            Marshal.GetDelegateForFunctionPointer<StopPlaybackDel>((IntPtr)_vt[27])(_ptr, stop, out _, scale);
            return 0;
        }

        public int GetFrameBytes(IntPtr frame, out IntPtr bytes, Action<string>? log)
        {
            bytes = IntPtr.Zero;
            Guid g = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qhr = Marshal.QueryInterface(frame, ref g, out IntPtr buf);
            if (qhr != 0) { log?.Invoke($"QI failed: 0x{qhr:X8}"); return qhr; }
            void** bvt = *(void***)buf;
            Marshal.GetDelegateForFunctionPointer<StartAccessDel>((IntPtr)bvt[4])(buf, 2);
            int gbHr = Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[3])(buf, out bytes);
            _bufPtr = buf; _bufVt = bvt;
            return gbHr;
        }

        private IntPtr _bufPtr; private void** _bufVt;

        public void ReleaseFrame(IntPtr frame)
        {
            if (frame == IntPtr.Zero) return;
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
