using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using DeckLinkAPI;
using SharpDX;
using SharpDX.Direct3D11;
using SharpDX.DXGI;

namespace MonitorToDeckLink
{
    public class MonitorInfo
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public string DeviceName { get; set; } = "";
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsPrimary { get; set; }
        public override string ToString() =>
            $"{(IsPrimary ? "★ " : "")}{Name}  ({Width}×{Height})  [{DeviceName}]";
    }

    public class DeckLinkDeviceInfo
    {
        public string Name { get; set; } = "";
        public IDeckLink Device { get; set; } = null!;
        public override string ToString() => Name;
    }

    public class OutputFormatInfo
    {
        public string Label { get; set; } = "";
        public _BMDDisplayMode Mode { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public double FrameRate { get; set; }
        public override string ToString() => Label;
    }

    public partial class MainWindow : Window
    {
        private CancellationTokenSource? _cts;
        private Task? _captureTask;

        private readonly List<OutputFormatInfo> _formats = new()
        {
            new() { Label = "1080p 23.98",  Mode = _BMDDisplayMode.bmdModeHD1080p2398, Width = 1920, Height = 1080, FrameRate = 24000.0/1001 },
            new() { Label = "1080p 24",     Mode = _BMDDisplayMode.bmdModeHD1080p24,   Width = 1920, Height = 1080, FrameRate = 24 },
            new() { Label = "1080p 25",     Mode = _BMDDisplayMode.bmdModeHD1080p25,   Width = 1920, Height = 1080, FrameRate = 25 },
            new() { Label = "1080p 29.97",  Mode = _BMDDisplayMode.bmdModeHD1080p2997, Width = 1920, Height = 1080, FrameRate = 30000.0/1001 },
            new() { Label = "1080p 30",     Mode = _BMDDisplayMode.bmdModeHD1080p30,   Width = 1920, Height = 1080, FrameRate = 30 },
            new() { Label = "1080i 50",     Mode = _BMDDisplayMode.bmdModeHD1080i50,   Width = 1920, Height = 1080, FrameRate = 25 },
            new() { Label = "1080i 59.94",  Mode = _BMDDisplayMode.bmdModeHD1080i5994, Width = 1920, Height = 1080, FrameRate = 30000.0/1001 },
            new() { Label = "1080i 60",     Mode = _BMDDisplayMode.bmdModeHD1080i6000, Width = 1920, Height = 1080, FrameRate = 30 },
            new() { Label = "1080p 50",     Mode = _BMDDisplayMode.bmdModeHD1080p50,   Width = 1920, Height = 1080, FrameRate = 50 },
            new() { Label = "1080p 59.94",  Mode = _BMDDisplayMode.bmdModeHD1080p5994, Width = 1920, Height = 1080, FrameRate = 60000.0/1001 },
            new() { Label = "1080p 60",     Mode = _BMDDisplayMode.bmdModeHD1080p6000, Width = 1920, Height = 1080, FrameRate = 60 },
            new() { Label = "720p 50",      Mode = _BMDDisplayMode.bmdModeHD720p50,    Width = 1280, Height = 720,  FrameRate = 50 },
            new() { Label = "720p 59.94",   Mode = _BMDDisplayMode.bmdModeHD720p5994,  Width = 1280, Height = 720,  FrameRate = 60000.0/1001 },
            new() { Label = "720p 60",      Mode = _BMDDisplayMode.bmdModeHD720p60,    Width = 1280, Height = 720,  FrameRate = 60 },
            new() { Label = "2160p 23.98",  Mode = _BMDDisplayMode.bmdMode4K2160p2398, Width = 3840, Height = 2160, FrameRate = 24000.0/1001 },
            new() { Label = "2160p 25",     Mode = _BMDDisplayMode.bmdMode4K2160p25,   Width = 3840, Height = 2160, FrameRate = 25 },
            new() { Label = "2160p 29.97",  Mode = _BMDDisplayMode.bmdMode4K2160p2997, Width = 3840, Height = 2160, FrameRate = 30000.0/1001 },
            new() { Label = "2160p 30",     Mode = _BMDDisplayMode.bmdMode4K2160p30,   Width = 3840, Height = 2160, FrameRate = 30 },
        };

        public MainWindow() { InitializeComponent(); }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            PopulateMonitors();
            PopulateDeckLinks();
            PopulateFormats();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => StopCapture();

        private void PopulateMonitors()
        {
            var monitors = new List<MonitorInfo>();
            try
            {
                using var factory = new Factory1();
                int adapterIdx = 0;
                foreach (var adapter in factory.Adapters1)
                {
                    int outIdx = 0;
                    foreach (var dxgiOut in adapter.Outputs)
                    {
                        var desc = dxgiOut.Description;
                        monitors.Add(new MonitorInfo
                        {
                            Index = monitors.Count,
                            Name = $"Monitor {monitors.Count + 1}",
                            DeviceName = desc.DeviceName,
                            Width  = desc.DesktopBounds.Right - desc.DesktopBounds.Left,
                            Height = desc.DesktopBounds.Bottom - desc.DesktopBounds.Top,
                            IsPrimary = adapterIdx == 0 && outIdx == 0
                        });
                        dxgiOut.Dispose();
                        outIdx++;
                    }
                    adapter.Dispose();
                    adapterIdx++;
                }
            }
            catch (Exception ex) { SetStatus($"Monitor error: {ex.Message}", true); }

            cmbMonitors.ItemsSource = monitors;
            if (monitors.Count > 0) cmbMonitors.SelectedIndex = 0;
        }

        private void PopulateDeckLinks()
        {
            var devices = new List<DeckLinkDeviceInfo>();
            try
            {
                var iterator = new CDeckLinkIterator() as IDeckLinkIterator
                    ?? throw new Exception("Cannot create DeckLink iterator.");
                while (true)
                {
                    iterator.Next(out IDeckLink device);
                    if (device == null) break;
                    device.GetDisplayName(out string name);
                    devices.Add(new DeckLinkDeviceInfo { Name = name, Device = device });
                }
            }
            catch (Exception ex) { SetStatus($"DeckLink error: {ex.Message}", true); }

            cmbDeckLinks.ItemsSource = devices;
            if (devices.Count > 0) cmbDeckLinks.SelectedIndex = 0;
        }

        private void PopulateFormats()
        {
            cmbFormats.ItemsSource = _formats;
            cmbFormats.SelectedIndex = 2;
        }

        private void cmbMonitors_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
        private void cmbDeckLinks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMonitors.SelectedItem is not MonitorInfo monitor)        { SetStatus("Select a monitor.", true); return; }
            if (cmbDeckLinks.SelectedItem is not DeckLinkDeviceInfo dlInfo) { SetStatus("Select a DeckLink device.", true); return; }
            if (cmbFormats.SelectedItem is not OutputFormatInfo format)     { SetStatus("Select an output format.", true); return; }

            btnStart.IsEnabled = false;
            btnStop.IsEnabled  = true;
            SetStatus($"Starting: {monitor.Name} → {dlInfo.Name} @ {format.Label}");

            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            _captureTask = Task.Run(() => CaptureLoop(monitor, dlInfo.Device, format, token), token);
            _captureTask.ContinueWith(t => Dispatcher.Invoke(() =>
            {
                btnStart.IsEnabled = true;
                btnStop.IsEnabled  = false;
                SetStatus(t.IsFaulted ? $"Error: {t.Exception?.InnerException?.Message}" : "Stopped.", t.IsFaulted);
            }));
        }

        private void btnStop_Click(object sender, RoutedEventArgs e) => StopCapture();

        private void StopCapture() { _cts?.Cancel(); _captureTask?.Wait(3000); _cts = null; }

        private void SetStatus(string msg, bool isError = false) => Dispatcher.Invoke(() =>
        {
            txtStatus.Text = msg;
            txtStatus.Foreground = isError
                ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(243, 139, 168))
                : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(166, 227, 161));
        });

        // ── Capture + output loop ─────────────────────────────────────────────

        private unsafe void CaptureLoop(MonitorInfo monitor, IDeckLink deckLink,
            OutputFormatInfo format, CancellationToken ct)
        {
            // D3D11 device
            using var d3dDevice = new SharpDX.Direct3D11.Device(
                SharpDX.Direct3D.DriverType.Hardware, DeviceCreationFlags.BgraSupport);
            using var dxgiDevice = d3dDevice.QueryInterface<SharpDX.DXGI.Device>();
            using var adapter    = dxgiDevice.GetParent<Adapter>();
            using var factory    = adapter.GetParent<Factory1>();

            // Find the matching DXGI output for the selected monitor
            Output1? dupeOutput = null;
            foreach (var adpt in factory.Adapters1)
            {
                foreach (var dxgiOut in adpt.Outputs)
                {
                    if (dxgiOut.Description.DeviceName == monitor.DeviceName)
                    {
                        dupeOutput = dxgiOut.QueryInterface<Output1>();
                        dxgiOut.Dispose();
                        goto foundOutput;
                    }
                    dxgiOut.Dispose();
                }
                adpt.Dispose();
            }
            foundOutput:
            if (dupeOutput == null)
                throw new Exception($"DXGI output not found: {monitor.DeviceName}");

            using var deskDupe = dupeOutput.DuplicateOutput(d3dDevice);
            dupeOutput.Dispose();

            // Staging texture for CPU readback
            using var stagingTex = new Texture2D(d3dDevice, new Texture2DDescription
            {
                Width = monitor.Width, Height = monitor.Height,
                MipLevels = 1, ArraySize = 1,
                Format = Format.B8G8R8A8_UNorm,
                SampleDescription = new SampleDescription(1, 0),
                Usage = ResourceUsage.Staging,
                BindFlags = BindFlags.None,
                CpuAccessFlags = CpuAccessFlags.Read,
                OptionFlags = ResourceOptionFlags.None
            });

            // DeckLink output
            var deckOutput = (IDeckLinkOutput)deckLink;
            deckOutput.EnableVideoOutput(format.Mode, _BMDVideoOutputFlags.bmdVideoOutputFlagDefault);

            long   frameDuration  = (long)(TimeSpan.TicksPerSecond / format.FrameRate);
            double framePeriodMs  = 1000.0 / format.FrameRate;
            var    sw             = Stopwatch.StartNew();
            long   frameNumber    = 0;

            // Keep last good frame data so we can re-send on timeout
            byte[]? lastFrame = null;

            while (!ct.IsCancellationRequested)
            {
                // Frame timing
                double targetMs = frameNumber * framePeriodMs;
                double waitMs   = targetMs - sw.Elapsed.TotalMilliseconds;
                if (waitMs > 1) Thread.Sleep((int)waitMs - 1);
                while (sw.Elapsed.TotalMilliseconds < targetMs) { /* spin */ }

                // Try to get a new desktop frame
                bool gotNewFrame = false;
                try
                {
                    int hr = deskDupe.TryAcquireNextFrame(0,
                        out OutputDuplicateFrameInformation frameInfo,
                        out SharpDX.DXGI.Resource desktopResource);

                    if (hr >= 0 && desktopResource != null)
                    {
                        using (desktopResource)
                        using (var desktopTex = desktopResource.QueryInterface<Texture2D>())
                            d3dDevice.ImmediateContext.CopyResource(desktopTex, stagingTex);

                        deskDupe.ReleaseFrame();
                        gotNewFrame = true;
                    }
                }
                catch (SharpDX.SharpDXException) { /* timeout or mode change - use last frame */ }

                // Build UYVY frame
                int    deckRowBytes = format.Width * 2;
                byte[] uyvyBuf      = new byte[deckRowBytes * format.Height];

                if (gotNewFrame)
                {
                    var mapped = d3dDevice.ImmediateContext.MapSubresource(
                        stagingTex, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                    try
                    {
                        fixed (byte* dst = uyvyBuf)
                            BgraToUyvy((byte*)mapped.DataPointer, mapped.RowPitch,
                                monitor.Width, monitor.Height,
                                dst, format.Width, format.Height);
                    }
                    finally { d3dDevice.ImmediateContext.UnmapSubresource(stagingTex, 0); }

                    lastFrame = uyvyBuf;
                }
                else if (lastFrame != null)
                {
                    uyvyBuf = lastFrame; // repeat last frame
                }
                else
                {
                    frameNumber++;
                    continue; // no frame yet at all
                }

                // Send to DeckLink
                deckOutput.CreateVideoFrame(
                    format.Width, format.Height, deckRowBytes,
                    _BMDPixelFormat.bmdFormat8BitYUV,
                    _BMDFrameFlags.bmdFrameFlagDefault,
                    out IDeckLinkMutableVideoFrame deckFrame);

                // Copy UYVY bytes into DeckLink frame via Marshal
                deckFrame.GetBytes(out IntPtr deckPtr);
                fixed (byte* src = uyvyBuf)
                    Buffer.MemoryCopy(src, (void*)deckPtr, uyvyBuf.Length, uyvyBuf.Length);

                deckOutput.ScheduleVideoFrame(deckFrame,
                    frameNumber * frameDuration, frameDuration, TimeSpan.TicksPerSecond);
                Marshal.ReleaseComObject(deckFrame);

                if (frameNumber == 0)
                    deckOutput.StartScheduledPlayback(0, TimeSpan.TicksPerSecond, 1.0);

                frameNumber++;

                if (frameNumber % (long)format.FrameRate == 0)
                {
                    deckOutput.GetBufferedVideoFrameCount(out uint buffered);
                    Dispatcher.Invoke(() =>
                        txtStatus.Text = $"Running — frame {frameNumber}  buffered: {buffered}");
                }
            }

            deckOutput.StopScheduledPlayback(0, out _, TimeSpan.TicksPerSecond);
            deckOutput.DisableVideoOutput();
        }

        // BT.709 BGRA → UYVY 4:2:2 with scaling
        private static unsafe void BgraToUyvy(
            byte* src, int srcPitch, int srcW, int srcH,
            byte* dst, int dstW, int dstH)
        {
            float scaleX = (float)srcW / dstW;
            float scaleY = (float)srcH / dstH;
            int   dstPitch = dstW * 2;

            for (int y = 0; y < dstH; y++)
            {
                int   srcY   = Math.Min((int)(y * scaleY), srcH - 1);
                byte* srcRow = src + srcY * srcPitch;
                byte* dstRow = dst + y   * dstPitch;

                for (int x = 0; x < dstW; x += 2)
                {
                    byte* p0 = srcRow + Math.Min((int)(x       * scaleX), srcW - 1) * 4;
                    byte* p1 = srcRow + Math.Min((int)((x + 1) * scaleX), srcW - 1) * 4;

                    float y0 = 16 + 0.1826f * p0[2] + 0.6142f * p0[1] + 0.0620f * p0[0];
                    float y1 = 16 + 0.1826f * p1[2] + 0.6142f * p1[1] + 0.0620f * p1[0];
                    float cb = 128 - 0.1006f * p0[2] - 0.3386f * p0[1] + 0.4392f * p0[0];
                    float cr = 128 + 0.4392f * p0[2] - 0.3989f * p0[1] - 0.0403f * p0[0];

                    dstRow[0] = (byte)Math.Clamp((int)cb, 0, 255);
                    dstRow[1] = (byte)Math.Clamp((int)y0, 16, 235);
                    dstRow[2] = (byte)Math.Clamp((int)cr, 0, 255);
                    dstRow[3] = (byte)Math.Clamp((int)y1, 16, 235);
                    dstRow += 4;
                }
            }
        }
    }
}
