using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using DeckLinkAPI;
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

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            PopulateMonitors();
            PopulateDeckLinks();
            PopulateFormats();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopCapture();
        }

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
                    foreach (var dxgiOutput in adapter.Outputs)
                    {
                        var desc = dxgiOutput.Description;
                        monitors.Add(new MonitorInfo
                        {
                            Index = monitors.Count,
                            Name = $"Monitor {monitors.Count + 1}",
                            DeviceName = desc.DeviceName,
                            Width = desc.DesktopBounds.Right - desc.DesktopBounds.Left,
                            Height = desc.DesktopBounds.Bottom - desc.DesktopBounds.Top,
                            IsPrimary = adapterIdx == 0 && outIdx == 0
                        });
                        dxgiOutput.Dispose();
                        outIdx++;
                    }
                    adapter.Dispose();
                    adapterIdx++;
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Error enumerating monitors: {ex.Message}", isError: true);
            }

            cmbMonitors.ItemsSource = monitors;
            if (monitors.Count > 0) cmbMonitors.SelectedIndex = 0;
        }

        private void PopulateDeckLinks()
        {
            var devices = new List<DeckLinkDeviceInfo>();
            try
            {
                var iterator = new CDeckLinkIterator() as IDeckLinkIterator
                    ?? throw new Exception("Cannot create DeckLink iterator. Is Desktop Video installed?");

                while (true)
                {
                    iterator.Next(out IDeckLink device);
                    if (device == null) break;
                    device.GetDisplayName(out string name);
                    devices.Add(new DeckLinkDeviceInfo { Name = name, Device = device });
                }
            }
            catch (Exception ex)
            {
                SetStatus($"DeckLink error: {ex.Message}", isError: true);
            }

            cmbDeckLinks.ItemsSource = devices;
            if (devices.Count > 0) cmbDeckLinks.SelectedIndex = 0;
        }

        private void PopulateFormats()
        {
            cmbFormats.ItemsSource = _formats;
            cmbFormats.SelectedIndex = 2; // 1080p25 default
        }

        private void cmbMonitors_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
        private void cmbDeckLinks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMonitors.SelectedItem is not MonitorInfo monitor)
            { SetStatus("Please select a monitor.", isError: true); return; }
            if (cmbDeckLinks.SelectedItem is not DeckLinkDeviceInfo deckLinkInfo)
            { SetStatus("Please select a DeckLink device.", isError: true); return; }
            if (cmbFormats.SelectedItem is not OutputFormatInfo format)
            { SetStatus("Please select an output format.", isError: true); return; }

            btnStart.IsEnabled = false;
            btnStop.IsEnabled = true;
            SetStatus($"Starting: {monitor.Name} → {deckLinkInfo.Name} @ {format.Label}");

            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            _captureTask = Task.Run(() => CaptureLoop(monitor, deckLinkInfo.Device, format, token), token);
            _captureTask.ContinueWith(t =>
            {
                Dispatcher.Invoke(() =>
                {
                    btnStart.IsEnabled = true;
                    btnStop.IsEnabled = false;
                    if (t.IsFaulted)
                        SetStatus($"Error: {t.Exception?.InnerException?.Message}", isError: true);
                    else
                        SetStatus("Stopped.");
                });
            });
        }

        private void btnStop_Click(object sender, RoutedEventArgs e) => StopCapture();

        private void StopCapture()
        {
            _cts?.Cancel();
            _captureTask?.Wait(3000);
            _cts = null;
        }

        private void SetStatus(string msg, bool isError = false)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = msg;
                txtStatus.Foreground = isError
                    ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(243, 139, 168))
                    : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(166, 227, 161));
            });
        }

        private unsafe void CaptureLoop(MonitorInfo monitor, IDeckLink deckLink,
            OutputFormatInfo format, CancellationToken ct)
        {
            // ── D3D11 device ──────────────────────────────────────────────────
            using var d3dDevice = new SharpDX.Direct3D11.Device(
                SharpDX.Direct3D.DriverType.Hardware,
                DeviceCreationFlags.BgraSupport);

            using var dxgiDevice = d3dDevice.QueryInterface<SharpDX.DXGI.Device>();
            using var adapter   = dxgiDevice.GetParent<Adapter>();
            using var factory   = adapter.GetParent<Factory1>();

            // ── Find the right DXGI output ────────────────────────────────────
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
                throw new Exception($"Could not find DXGI output for {monitor.DeviceName}");

            using var deskDupe = dupeOutput.DuplicateOutput(d3dDevice);
            dupeOutput.Dispose();

            // ── Staging texture for CPU readback ──────────────────────────────
            var stagingDesc = new Texture2DDescription
            {
                Width = monitor.Width, Height = monitor.Height,
                MipLevels = 1, ArraySize = 1,
                Format = Format.B8G8R8A8_UNorm,
                SampleDescription = new SampleDescription(1, 0),
                Usage = ResourceUsage.Staging,
                BindFlags = BindFlags.None,
                CpuAccessFlags = CpuAccessFlags.Read,
                OptionFlags = ResourceOptionFlags.None
            };
            using var stagingTex = new Texture2D(d3dDevice, stagingDesc);

            // ── DeckLink output ───────────────────────────────────────────────
            var deckOutput = (IDeckLinkOutput)deckLink;
            deckOutput.EnableVideoOutput(format.Mode,
                _BMDVideoOutputFlags.bmdVideoOutputFlagDefault);

            long frameDuration  = (long)(TimeSpan.TicksPerSecond / format.FrameRate);
            double framePeriodMs = 1000.0 / format.FrameRate;
            var sw = Stopwatch.StartNew();
            long frameNumber = 0;

            while (!ct.IsCancellationRequested)
            {
                // ── Timing ────────────────────────────────────────────────────
                double targetMs = frameNumber * framePeriodMs;
                double waitMs   = targetMs - sw.Elapsed.TotalMilliseconds;
                if (waitMs > 1) Thread.Sleep((int)waitMs - 1);
                while (sw.Elapsed.TotalMilliseconds < targetMs) { /* spin */ }

                // ── Capture frame ─────────────────────────────────────────────
                SharpDX.DXGI.Resource? desktopResource = null;
                try
                {
                    deskDupe.TryAcquireNextFrame(16,
                        out SharpDX.DXGI.OutduplFrameInfo frameInfo,
                        out desktopResource);
                }
                catch (SharpDXException ex) when (ex.ResultCode == SharpDX.DXGI.ResultCode.WaitTimeout)
                {
                    frameNumber++;
                    continue;
                }

                if (desktopResource == null) { frameNumber++; continue; }

                using (desktopResource)
                using (var desktopTex = desktopResource.QueryInterface<Texture2D>())
                {
                    d3dDevice.ImmediateContext.CopyResource(desktopTex, stagingTex);
                }
                deskDupe.ReleaseFrame();

                // ── Map + convert + send ──────────────────────────────────────
                var mapped = d3dDevice.ImmediateContext.MapSubresource(
                    stagingTex, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);

                try
                {
                    deckOutput.CreateVideoFrame(
                        format.Width, format.Height,
                        format.Width * 2,
                        _BMDPixelFormat.bmdFormat8BitYUV,
                        _BMDFrameFlags.bmdFrameFlagDefault,
                        out IDeckLinkMutableVideoFrame deckFrame);

                    deckFrame.GetBytes(out IntPtr deckPtr);

                    BgraToUyvy(
                        (byte*)mapped.DataPointer,
                        mapped.RowPitch,
                        monitor.Width, monitor.Height,
                        (byte*)deckPtr,
                        format.Width, format.Height);

                    deckOutput.ScheduleVideoFrame(deckFrame,
                        frameNumber * frameDuration,
                        frameDuration,
                        TimeSpan.TicksPerSecond);

                    Marshal.ReleaseComObject(deckFrame);
                }
                finally
                {
                    d3dDevice.ImmediateContext.UnmapSubresource(stagingTex, 0);
                }

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

        private static unsafe void BgraToUyvy(
            byte* src, int srcPitch, int srcW, int srcH,
            byte* dst, int dstW, int dstH)
        {
            float scaleX = (float)srcW / dstW;
            float scaleY = (float)srcH / dstH;
            int dstPitch = dstW * 2;

            for (int y = 0; y < dstH; y++)
            {
                int srcY   = Math.Min((int)(y * scaleY), srcH - 1);
                byte* srcRow = src + srcY * srcPitch;
                byte* dstRow = dst + y  * dstPitch;

                for (int x = 0; x < dstW; x += 2)
                {
                    int srcX0 = Math.Min((int)(x * scaleX), srcW - 1);
                    int srcX1 = Math.Min((int)((x + 1) * scaleX), srcW - 1);

                    byte* p0 = srcRow + srcX0 * 4;
                    byte* p1 = srcRow + srcX1 * 4;

                    // BT.709 BGRA → YCbCr
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
