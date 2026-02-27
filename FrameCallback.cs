using System;
using System.Runtime.InteropServices;

namespace MonitorToDeckLink
{
    // Minimal IDeckLinkVideoOutputCallback implementation
    // Required before ScheduleVideoFrame will work
    // Interface GUID from OBS SDK: IDeckLinkVideoOutputCallback
    [ComVisible(true)]
    [Guid("178A3407-A3A8-4B16-9FEB-5B127CD8434D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkVideoOutputCallback
    {
        [PreserveSig] int ScheduledFrameCompleted(IntPtr completedFrame, int result);
        [PreserveSig] int ScheduledPlaybackHasStopped();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class FrameCallback : IDeckLinkVideoOutputCallback
    {
        public volatile int CompletedFrames = 0;

        public int ScheduledFrameCompleted(IntPtr completedFrame, int result)
        {
            CompletedFrames++;
            return 0; // S_OK
        }

        public int ScheduledPlaybackHasStopped() => 0;
    }
}
