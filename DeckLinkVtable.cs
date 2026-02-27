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

        // Methods are Offset + 3 (for IUnknown)
        // EnableVideoOutput = Method 4 -> Slot 7
        public int EnableVideoOutput(int mode, int flags) =>
            Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vt[7])(_ptr, mode, flags);

        public int DisableVideoOutput() =>
            Marshal.GetDelegateForFunctionPointer<DisableVideoOutputDel>((IntPtr)_vt[8])(_ptr);

        // CreateVideoFrame = Method 6 -> Slot 9
        public int CreateVideoFrame(int w, int h, int rb, int fmt, int flags, out IntPtr frame) =>
            Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vt[9])(_ptr, w, h, rb, fmt, flags, out frame);

        // ScheduleVideoFrame = Method 11 -> Slot 14
        public int ScheduleVideoFrame(IntPtr frame, long time, long dur, long scale) =>
            Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vt[14])(_ptr, frame, time, dur, scale);

        // SetScheduledFrameCompletionCallback = Method 12 -> Slot 15
        public int SetFrameCallback(IntPtr cb) =>
            Marshal.GetDelegateForFunctionPointer<SetCallbackDel>((IntPtr)_vt[15])(_ptr, cb);

        public int GetBufferedVideoFrameCount(out uint count) =>
            Marshal.GetDelegateForFunctionPointer<GetBufferedCountDel>((IntPtr)_vt[16])(_ptr, out count);

        public int StartScheduledPlayback(long start, long scale, double speed) =>
            Marshal.GetDelegateForFunctionPointer<StartPlaybackDel>((IntPtr)_vt[26])(_ptr, start, scale, speed);

        public int GetFrameBytes(IntPtr frame, out IntPtr bytes)
        {
            bytes = IntPtr.Zero;
            Guid g = new Guid("CF9EB134-0374-4C5B-95FA-1EC14819FF62"); // IDeckLinkVideoFrame
            int qhr = Marshal.QueryInterface(frame, ref g, out IntPtr buf);
            if (qhr != 0) return qhr;
            
            void** bvt = *(void***)buf;
            // GetBytes = Method 5 -> Slot 8 (Inside VideoFrame Vtable)
            int gbHr = Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)bvt[8])(buf, out bytes);
            return gbHr;
        }

        public void Dispose()
        {
            if (_ptr == IntPtr.Zero) return;
            Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_vt[2])(_ptr);
            _ptr = IntPtr.Zero;
        }
    }
}
