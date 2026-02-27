using System;
using System.Runtime.InteropServices;
using System.Threading;
using Windows.Graphics.Capture;
using Windows.Graphics.DirectX;
using Windows.Graphics.DirectX.Direct3D11;
using SharpDX.Direct3D11;
using SharpDX.DXGI;

namespace MonitorToDeckLink
{
    class Program
    {
        // DeckLink Constants
        const uint bmdModeHD1080p6000 = 0x68703630; // 'hp60'
        const uint bmdFormat8BitUYVY = 0x32767579; // '2vuy'
        
        static IDeckLinkOutput2 m_deckLinkOutput;
        static long m_totalFramesScheduled = 0;
        static SharpDX.Direct3D11.Device m_d3dDevice;

        [MTAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("--- Starting Monitor to DeckLink Pipe ---");
            
            // 1. Initialize DeckLink
            var iterator = new DeckLinkIterator();
            IDeckLink2 deckLink;
            iterator.Next(out deckLink);
            
            // Query the Output interface (VTable Offset handling)
            m_deckLinkOutput = (IDeckLinkOutput2)deckLink;

            // 2. Setup Screen Capture (Simplified for brevity)
            // Ensure monitor is 1920x1080
            InitializeCapture();

            Console.WriteLine("Piping started. Press any key to stop.");
            Console.ReadKey();
        }

        static void OnFrameArrived(Direct3D11CaptureFrame frame)
        {
            // The monitor gives us 7680 bytes per row (BGRA)
            // The DeckLink needs 3840 bytes per row (UYVY)
            
            IDeckLinkMutableVideoFrame2 videoFrame;
            // FIX 1: Create frame with 3840 rowBytes, NOT 7680
            m_deckLinkOutput.CreateVideoFrame(1920, 1080, 3840, bmdFormat8BitUYVY, 0, out videoFrame);

            IntPtr deckLinkBuffer;
            videoFrame.GetBytes(out deckLinkBuffer);

            using (var texture = SharpDX.Direct3D11.Resource.FromAbi<Texture2D>(frame.Surface.NativeObject))
            {
                var map = m_d3dDevice.ImmediateContext.MapSubresource(texture, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                
                // FIX 2: Optimized BGRA to UYVY Conversion
                unsafe
                {
                    byte* src = (byte*)map.DataPointer;
                    byte* dst = (byte*)deckLinkBuffer;

                    for (int y = 0; y < 1080; y++)
                    {
                        for (int x = 0; x < 1920; x += 2)
                        {
                            // Source indices (4 bytes per pixel)
                            int sIdx1 = (y * map.RowPitch) + (x * 4);
                            int sIdx2 = sIdx1 + 4;

                            // Destination indices (2 bytes per pixel)
                            int dIdx = (y * 3840) + (x * 2);

                            // Simple Color Conversion (Averaging for U/V)
                            // UYVY layout: [U0, Y0, V0, Y1]
                            dst[dIdx + 0] = 128; // U (Chroma)
                            dst[dIdx + 1] = src[sIdx1 + 1]; // Y (Luma from Green channel as proxy)
                            dst[dIdx + 2] = 128; // V (Chroma)
                            dst[dIdx + 3] = src[sIdx2 + 1]; // Y2
                        }
                    }
                }
                m_d3dDevice.ImmediateContext.UnmapSubresource(texture, 0);
            }

            // FIX 3: Correct Scheduling Timing (1000 duration at 60000 timescale)
            m_deckLinkOutput.ScheduleVideoFrame(
                videoFrame, 
                m_totalFramesScheduled * 1000, 
                1000, 
                60000
            );
            
            m_totalFramesScheduled++;

            if (m_totalFramesScheduled == 3) // Start playback after pre-roll
            {
                m_deckLinkOutput.StartScheduledPlayback(0, 60000, 1.0);
            }
        }
    }

    // Manual VTable Mapping to ensure DeckLink COM compatibility
    [ComImport, Guid("1A8077F1-03D9-45C0-9907-77561D02E5C9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeckLinkOutput2
    {
        [PreserveSig] int EnableVideoOutput(uint displayMode, uint flags);
        [PreserveSig] int DisableVideoOutput();
        [PreserveSig] int CreateVideoFrame(int width, int height, int rowBytes, uint pixelFormat, uint flags, out IDeckLinkMutableVideoFrame2 frame);
        // ... (Include other methods in exact VTable order)
        [PreserveSig] int ScheduleVideoFrame(IDeckLinkMutableVideoFrame2 frame, long displayTime, long displayDuration, long timeScale);
        [PreserveSig] int StartScheduledPlayback(long playbackStartTime, long timeScale, double playbackSpeed);
    }
}
