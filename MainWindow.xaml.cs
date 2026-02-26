using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
        public IntPtr RawPtr { get; set; } = IntPtr.Zero;  // Raw COM pointer for QI
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
        private readonly string _logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "MonitorToDeckLink.log");

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
            Log($"=== MonitorToDeckLink started {DateTime.Now} ===");
            Log($"Log file: {_logPath}");
            PopulateMonitors();
            PopulateDeckLinks();
            PopulateFormats();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => StopCapture();

        private void btnClearLog_Click(object sender, RoutedEventArgs e) => txtLog.Text = "";
        private void btnCopyLog_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(txtLog.Text);
            SetStatus("Log copied to clipboard.");
        }

        // ── Logging ───────────────────────────────────────────────────────────

        private void Log(string msg)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
            try { File.AppendAllText(_logPath, line + Environment.NewLine); } catch { }
            Dispatcher.Invoke(() =>
            {
                txtLog.Text += line + "\n";
                logScroller.ScrollToEnd();
            });
        }

        // ── Populate monitors ─────────────────────────────────────────────────

        private void PopulateMonitors()
        {
            var monitors = new List<MonitorInfo>();
            try
            {
                Log("Enumerating monitors via DXGI...");
                using var factory = new Factory1();
                int adapterIdx = 0;
                foreach (var adapter in factory.Adapters1)
                {
                    Log($"  Adapter {adapterIdx}: {adapter.Description.Description}");
                    int outIdx = 0;
                    foreach (var dxgiOut in adapter.Outputs)
                    {
                        var desc = dxgiOut.Description;
                        var m = new MonitorInfo
                        {
                            Index = monitors.Count,
                            Name = $"Monitor {monitors.Count + 1}",
                            DeviceName = desc.DeviceName,
                            Width  = desc.DesktopBounds.Right - desc.DesktopBounds.Left,
                            Height = desc.DesktopBounds.Bottom - desc.DesktopBounds.Top,
                            IsPrimary = adapterIdx == 0 && outIdx == 0
                        };
                        monitors.Add(m);
                        Log($"    Output {outIdx}: {desc.DeviceName} {m.Width}x{m.Height} primary={m.IsPrimary}");
                        dxgiOut.Dispose();
                        outIdx++;
                    }
                    adapter.Dispose();
                    adapterIdx++;
                }
                Log($"Found {monitors.Count} monitor(s).");
            }
            catch (Exception ex) { Log($"ERROR enumerating monitors: {ex}"); SetStatus($"Monitor error: {ex.Message}", true); }

            cmbMonitors.ItemsSource = monitors;
            if (monitors.Count > 0) cmbMonitors.SelectedIndex = 0;
        }

        // ── Populate DeckLink devices ─────────────────────────────────────────

        private void PopulateDeckLinks()
        {
            var devices = new List<DeckLinkDeviceInfo>();
            try
            {
                Log("Enumerating DeckLink devices...");
                var iterator = new CDeckLinkIterator() as IDeckLinkIterator
                    ?? throw new Exception("Cannot create DeckLink iterator. Is Desktop Video installed?");
                while (true)
                {
                    iterator.Next(out IDeckLink device);
                    if (device == null) break;
                    device.GetDisplayName(out string name);
                    // Capture the raw COM pointer BEFORE the RCW takes over
                    IntPtr rawPtr = Marshal.GetIUnknownForObject(device);
                    // AddRef already happened in GetIUnknownForObject, keep one ref alive
                    devices.Add(new DeckLinkDeviceInfo { Name = name, Device = device, RawPtr = rawPtr });
                    Log($"  Found DeckLink device: {name}  rawPtr=0x{rawPtr:X}");
                }
                Log($"Found {devices.Count} DeckLink device(s).");
            }
            catch (Exception ex) { Log($"ERROR enumerating DeckLink: {ex}"); SetStatus($"DeckLink error: {ex.Message}", true); }

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

        // ── Start / Stop ──────────────────────────────────────────────────────

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMonitors.SelectedItem is not MonitorInfo monitor)        { SetStatus("Select a monitor.", true); return; }
            if (cmbDeckLinks.SelectedItem is not DeckLinkDeviceInfo dlInfo) { SetStatus("Select a DeckLink device.", true); return; }
            if (cmbFormats.SelectedItem is not OutputFormatInfo format)     { SetStatus("Select an output format.", true); return; }

            btnStart.IsEnabled = false;
            btnStop.IsEnabled  = true;
            SetStatus($"Starting: {monitor.Name} → {dlInfo.Name} @ {format.Label}");
            Log($"--- Starting capture: {monitor} → {dlInfo.Name} @ {format.Label} ---");

            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            _captureTask = Task.Run(() => CaptureLoop(monitor, dlInfo, format, token), token);
            _captureTask.ContinueWith(t => Dispatcher.Invoke(() =>
            {
                btnStart.IsEnabled = true;
                btnStop.IsEnabled  = false;
                if (t.IsFaulted)
                {
                    var msg = t.Exception?.InnerException?.ToString() ?? "Unknown error";
                    Log($"ERROR: {msg}");
                    SetStatus($"Error: {t.Exception?.InnerException?.Message}", true);
                }
                else
                {
                    Log("--- Capture stopped ---");
                    SetStatus("Stopped.");
                }
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

        private unsafe void CaptureLoop(MonitorInfo monitor, DeckLinkDeviceInfo deckLinkInfo,
            OutputFormatInfo format, CancellationToken ct)
        {
            IDeckLink deckLink = deckLinkInfo.Device;
            Log("Creating D3D11 device...");
            using var d3dDevice = new SharpDX.Direct3D11.Device(
                SharpDX.Direct3D.DriverType.Hardware, DeviceCreationFlags.BgraSupport);
            Log($"D3D11 device created: {d3dDevice.FeatureLevel}");

            using var dxgiDevice = d3dDevice.QueryInterface<SharpDX.DXGI.Device>();
            using var adapter    = dxgiDevice.GetParent<Adapter>();
            using var factory    = adapter.GetParent<Factory1>();

            Log($"Finding DXGI output for: {monitor.DeviceName}");
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
            if (dupeOutput == null) throw new Exception($"DXGI output not found: {monitor.DeviceName}");
            Log("DXGI output found. Creating desktop duplication...");

            using var deskDupe = dupeOutput.DuplicateOutput(d3dDevice);
            dupeOutput.Dispose();
            Log("Desktop duplication created.");

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
            Log($"Staging texture created: {monitor.Width}x{monitor.Height}");

            Log($"Enabling DeckLink video output: {format.Label} ({format.Width}x{format.Height})");

            // Try all known IDeckLinkOutput GUIDs across SDK versions
            var outputCandidates = new (string name, string guid)[]
            {
                ("IDeckLinkOutput",        "1A8077F1-9FE2-4533-8147-2294305E253F"),
                ("IDeckLinkOutput_v14_2_1","BE2D9020-461E-442F-84B7-E949CB953B9D"),
                ("IDeckLinkOutput_v11_4",  "065A0F6C-C508-4D0D-B919-F5EB0EBFC96B"),
                ("IDeckLinkOutput_v10_11", "CC5C8A6E-3F2F-4B3A-87EA-FD78AF300564"),
            };

            // Use the raw COM pointer stored at enumeration time
            IntPtr iunkPtr = deckLinkInfo.RawPtr != IntPtr.Zero
                ? deckLinkInfo.RawPtr
                : Marshal.GetIUnknownForObject(deckLink);
            Log($"IUnknown ptr: 0x{iunkPtr:X}");

            IntPtr outputPtr = IntPtr.Zero;
            string foundOutputName = "";
            foreach (var (name, guidStr) in outputCandidates)
            {
                Guid g = new Guid(guidStr);
                int hr = Marshal.QueryInterface(iunkPtr, ref g, out IntPtr ptr);
                Log($"  QI {name}: 0x{hr:X8}  ptr=0x{ptr:X}");
                if (hr == 0 && ptr != IntPtr.Zero)
                {
                    outputPtr = ptr;
                    foundOutputName = name;
                    break;
                }
            }
            // Don't release iunkPtr if it came from RawPtr (we hold that ref)
            if (deckLinkInfo.RawPtr == IntPtr.Zero) Marshal.Release(iunkPtr);

            if (outputPtr == IntPtr.Zero)
                throw new Exception("No IDeckLinkOutput interface found on this device. Check device selection.");

            Log($"Acquired {foundOutputName}.");
            var deckOutput = (IDeckLinkOutput)Marshal.GetObjectForIUnknown(outputPtr);
            Marshal.Release(outputPtr);
            Log("IDeckLinkOutput interface acquired.");

            deckOutput.EnableVideoOutput(format.Mode, _BMDVideoOutputFlags.bmdVideoOutputFlagDefault);
            Log("DeckLink video output enabled.");

            long   frameDuration  = (long)(TimeSpan.TicksPerSecond / format.FrameRate);
            double framePeriodMs  = 1000.0 / format.FrameRate;
            var    sw             = Stopwatch.StartNew();
            long   frameNumber    = 0;
            byte[]? lastFrame     = null;
            int    timeoutCount   = 0;

            Log("Entering capture loop...");

            while (!ct.IsCancellationRequested)
            {
                double targetMs = frameNumber * framePeriodMs;
                double waitMs   = targetMs - sw.Elapsed.TotalMilliseconds;
                if (waitMs > 1) Thread.Sleep((int)waitMs - 1);
                while (sw.Elapsed.TotalMilliseconds < targetMs) { }

                bool gotNewFrame = false;
                try
                {
                    SharpDX.Result hr = deskDupe.TryAcquireNextFrame(0,
                        out OutputDuplicateFrameInformation frameInfo,
                        out SharpDX.DXGI.Resource desktopResource);

                    if (hr.Success && desktopResource != null)
                    {
                        using (desktopResource)
                        using (var desktopTex = desktopResource.QueryInterface<Texture2D>())
                            d3dDevice.ImmediateContext.CopyResource(desktopTex, stagingTex);
                        deskDupe.ReleaseFrame();
                        gotNewFrame = true;
                        timeoutCount = 0;
                    }
                    else
                    {
                        timeoutCount++;
                    }
                }
                catch (SharpDX.SharpDXException ex)
                {
                    timeoutCount++;
                    if (timeoutCount == 1 || timeoutCount % 100 == 0)
                        Log($"Frame acquire exception (count={timeoutCount}): {ex.ResultCode} {ex.Message}");
                }

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
                    uyvyBuf = lastFrame;
                }
                else
                {
                    frameNumber++;
                    continue;
                }

                deckOutput.CreateVideoFrame(
                    format.Width, format.Height, deckRowBytes,
                    _BMDPixelFormat.bmdFormat8BitYUV,
                    _BMDFrameFlags.bmdFrameFlagDefault,
                    out IDeckLinkMutableVideoFrame deckFrame);

                var deckBuffer = (IDeckLinkVideoBuffer)deckFrame;
                deckBuffer.GetBytes(out IntPtr deckPtr);
                fixed (byte* src = uyvyBuf)
                    System.Buffer.MemoryCopy(src, (void*)deckPtr, uyvyBuf.Length, uyvyBuf.Length);

                deckOutput.ScheduleVideoFrame(deckFrame,
                    frameNumber * frameDuration, frameDuration, TimeSpan.TicksPerSecond);
                Marshal.ReleaseComObject(deckFrame);

                if (frameNumber == 0)
                {
                    Log("Starting scheduled playback...");
                    deckOutput.StartScheduledPlayback(0, TimeSpan.TicksPerSecond, 1.0);
                    Log("Scheduled playback started.");
                }

                frameNumber++;

                if (frameNumber % (long)format.FrameRate == 0)
                {
                    deckOutput.GetBufferedVideoFrameCount(out uint buffered);
                    string statusMsg = $"Running — frame {frameNumber}  buffered: {buffered}  timeouts: {timeoutCount}";
                    Dispatcher.Invoke(() => txtStatus.Text = statusMsg);
                    if (frameNumber % ((long)format.FrameRate * 5) == 0)
                        Log(statusMsg);
                }
            }

            Log("Stopping scheduled playback...");
            deckOutput.StopScheduledPlayback(0, out _, TimeSpan.TicksPerSecond);
            deckOutput.DisableVideoOutput();
            Log("DeckLink output disabled.");
        }

        // ── BT.709 BGRA → UYVY 4:2:2 ─────────────────────────────────────────

        private static unsafe void BgraToUyvy(
            byte* src, int srcPitch, int srcW, int srcH,
            byte* dst, int dstW, int dstH)
        {
            float scaleX  = (float)srcW / dstW;
            float scaleY  = (float)srcH / dstH;
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
