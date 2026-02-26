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

        public MainWindow()
        {
            InitializeComponent();
            // Catch any unhandled exceptions including native access violations
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                Log($"UNHANDLED: {e.ExceptionObject}");
            System.Windows.Application.Current.DispatcherUnhandledException += (s, e) =>
            {
                Log($"DISPATCHER UNHANDLED: {e.Exception}");
                e.Handled = true;
            };
            TaskScheduler.UnobservedTaskException += (s, e) =>
            {
                Log($"TASK UNHANDLED: {e.Exception}");
                e.SetObserved();
            };
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Log($"=== MonitorToDeckLink started {DateTime.Now} ===");
            Log($"Log file: {_logPath}");
            PopulateMonitors();
            PopulateDeckLinks();
            PopulateFormats();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => StopCapture();

        private IntPtr GetDeckLinkOutputPtr(IntPtr deckLinkRawPtr)
        {
            var candidates = new[] {
                "1A8077F1-9FE2-4533-8147-2294305E253F",  // IDeckLinkOutput
                "BE2D9020-461E-442F-84B7-E949CB953B9D",  // IDeckLinkOutput_v14_2_1
                "065A0F6C-C508-4D0D-B919-F5EB0EBFC96B",  // IDeckLinkOutput_v11_4
                "CC5C8A6E-3F2F-4B3A-87EA-FD78AF300564",  // IDeckLinkOutput_v10_11
            };
            foreach (var g in candidates)
            {
                Guid guid = new Guid(g);
                int hr = Marshal.QueryInterface(deckLinkRawPtr, ref guid, out IntPtr ptr);
                if (hr == 0 && ptr != IntPtr.Zero) return ptr;
            }
            return IntPtr.Zero;
        }

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
            // Capture the raw COM pointer on the UI thread (STA) before switching threads
            IntPtr rawOutputPtr = GetDeckLinkOutputPtr(dlInfo.RawPtr);
            Log($"Pre-captured IDeckLinkOutput ptr on UI thread: 0x{rawOutputPtr:X}");
            if (rawOutputPtr == IntPtr.Zero)
            {
                SetStatus("Failed to get DeckLink output interface.", true);
                btnStart.IsEnabled = true;
                btnStop.IsEnabled = false;
                return;
            }
            _captureTask = Task.Run(() => CaptureLoop(monitor, rawOutputPtr, format, token), token);
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

        private unsafe void CaptureLoop(MonitorInfo monitor, IntPtr outputPtr,
            OutputFormatInfo format, CancellationToken ct)
        {
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

            // outputPtr was pre-acquired on the UI thread and passed in
            Log("Setting up vtable caller...");
            using var deckOutput = new DeckLinkOutputVtable(outputPtr);
            Marshal.Release(outputPtr); // DeckLinkOutputVtable AddRefs internally
            Log($"Vtable entries:\n{deckOutput.VtableDump}");

            Log($"Display mode value: 0x{(int)format.Mode:X8}");
            // Try vtable slots 5-10 to find EnableVideoOutput
            // EnableVideoOutput signature: (this, BMDDisplayMode, BMDVideoOutputFlags) -> HRESULT
            int enableHr = -1;
            int enableSlot = -1;
            for (int slot = 5; slot <= 10; slot++)
            {
                Log($"Trying EnableVideoOutput at vtable[{slot}]...");
                try
                {
                    enableHr = deckOutput.CallEnableVideoOutput(slot, (int)format.Mode, 0);
                    Log($"  vtable[{slot}] returned 0x{enableHr:X8}");
                    // S_OK=0, and some DeckLink errors start with 0x89 or 0x8004
                    // A crash would throw; a wrong slot returns garbage HRESULT
                    // Accept if looks like a real HRESULT (top bit set = error, 0 = success)
                    if (enableHr == 0)
                    {
                        enableSlot = slot;
                        Log($"EnableVideoOutput succeeded at vtable[{slot}]");
                        break;
                    }
                    else if ((uint)enableHr >= 0x80000000 && (uint)enableHr <= 0x8FFFFFFF)
                    {
                        // Looks like a real HRESULT error - this is the right slot
                        enableSlot = slot;
                        Log($"EnableVideoOutput returned error at vtable[{slot}]: 0x{enableHr:X8}");
                        break;
                    }
                    // Otherwise garbage value - wrong slot, continue
                    Log($"  vtable[{slot}] returned non-HRESULT garbage, trying next slot");
                    // Must disable before trying another slot
                    try { deckOutput.CallDisableVideoOutput(slot + 1); } catch { }
                }
                catch (Exception ex)
                {
                    Log($"  vtable[{slot}] threw {ex.GetType().Name} - trying next");
                }
            }
            if (enableSlot < 0) throw new Exception("Could not find EnableVideoOutput in vtable slots 5-10");
            if (enableHr != 0) throw new Exception($"EnableVideoOutput failed: 0x{enableHr:X8}");
            Log($"DeckLink video output enabled via vtable[{enableSlot}].");

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
                    else { timeoutCount++; }
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
                else if (lastFrame != null) { uyvyBuf = lastFrame; }
                else { frameNumber++; continue; }

                // Allocate frame buffer, fill with UYVY, schedule
                IntPtr frameBytes = Marshal.AllocHGlobal(uyvyBuf.Length);
                try
                {
                    fixed (byte* src = uyvyBuf)
                        System.Buffer.MemoryCopy(src, (void*)frameBytes, uyvyBuf.Length, uyvyBuf.Length);

                    int schedHr = deckOutput.CreateAndScheduleFrame(
                        format.Width, format.Height, deckRowBytes,
                        frameBytes, uyvyBuf.Length,
                        frameNumber * frameDuration, frameDuration,
                        TimeSpan.TicksPerSecond);

                    if (schedHr != 0 && frameNumber < 5)
                        Log($"CreateAndScheduleFrame hr=0x{schedHr:X8} frame={frameNumber}");
                }
                finally { Marshal.FreeHGlobal(frameBytes); }

                if (frameNumber == 0)
                {
                    Log("Starting scheduled playback...");
                    int startHr = deckOutput.StartScheduledPlayback(0, TimeSpan.TicksPerSecond, 1.0);
                    Log($"StartScheduledPlayback hr=0x{startHr:X8}");
                }

                frameNumber++;

                if (frameNumber % (long)format.FrameRate == 0)
                {
                    uint buffered = deckOutput.GetBufferedVideoFrameCount();
                    string statusMsg = $"Running — frame {frameNumber}  buffered: {buffered}  timeouts: {timeoutCount}";
                    Dispatcher.Invoke(() => txtStatus.Text = statusMsg);
                    if (frameNumber % ((long)format.FrameRate * 5) == 0)
                        Log(statusMsg);
                }
            }

            Log("Stopping scheduled playback...");
            deckOutput.StopScheduledPlayback(0, TimeSpan.TicksPerSecond);
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

    // ── DeckLink Output vtable caller ─────────────────────────────────────────
    // Calls IDeckLinkOutput COM methods directly via function pointers.
    // IDeckLinkOutput vtable layout (from DeckLink SDK header):
    //   0: QueryInterface, 1: AddRef, 2: Release (IUnknown)
    //   3: DoesSupportVideoMode
    //   4: GetDisplayModeIterator
    //   5: SetScreenPreviewCallback
    //   6: EnableVideoOutput
    //   7: DisableVideoOutput
    //   8: SetVideoOutputFrameMemoryAllocator
    //   9: CreateVideoFrame
    //  10: SetAncillaryData
    //  11: SetVideoOutputConversionMode
    //  12: ScheduleVideoFrame
    //  13: GetBufferedVideoFrameCount
    //  14: StartScheduledPlayback
    //  15: StopScheduledPlayback
    //  16: IsScheduledPlaybackRunning
    //  17: GetScheduledStreamTime
    //  18: GetReferenceStatus
    //  19: EnableAudioOutput
    //  20: DisableAudioOutput
    //  21: WriteAudioSamplesSync
    //  22: BeginAudioPreroll
    //  23: EndAudioPreroll
    //  24: ScheduleAudioSamples
    //  25: GetBufferedAudioSampleFrameCount
    //  26: FlushBufferedAudioSamples
    //  27: SetAudioCallback
    //  28: GetOutputVideoFrameState

    internal unsafe class DeckLinkOutputVtable : IDisposable
    {
        private IntPtr _ptr;
        private void** _vtable;

        // Delegate types for each method we use
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int EnableVideoOutputDel(IntPtr self, int displayMode, int flags);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int DisableVideoOutputDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int CreateVideoFrameDel(IntPtr self, int width, int height, int rowBytes, int pixelFormat, int flags, out IntPtr outFrame);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int ScheduleVideoFrameDel(IntPtr self, IntPtr frame, long displayTime, long displayDuration, long timeScale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBufferedVideoFrameCountDel(IntPtr self, out uint buffered);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StartScheduledPlaybackDel(IntPtr self, long startTime, long timeScale, double speed);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int StopScheduledPlaybackDel(IntPtr self, long stopTime, out long actualStopTime, long timeScale);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetBytesDel(IntPtr self, out IntPtr buffer);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint AddRefDel(IntPtr self);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate uint ReleaseDel(IntPtr self);

        public string VtableDump { get; private set; } = "";

        public DeckLinkOutputVtable(IntPtr ptr)
        {
            _ptr = ptr;
            _vtable = *(void***)ptr;
            // Dump first 20 vtable entries for debugging
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < 20; i++)
                sb.Append($"  [{i}]=0x{(IntPtr)_vtable[i]:X}\n");
            VtableDump = sb.ToString();
            var addRef = Marshal.GetDelegateForFunctionPointer<AddRefDel>((IntPtr)_vtable[1]);
            addRef(_ptr);
        }

        public int CallEnableVideoOutput(int slot, int displayMode, int flags)
        {
            var fn = Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vtable[slot]);
            return fn(_ptr, displayMode, flags);
        }

        public int CallDisableVideoOutput(int slot)
        {
            var fn = Marshal.GetDelegateForFunctionPointer<DisableVideoOutputDel>((IntPtr)_vtable[slot]);
            return fn(_ptr);
        }

        public int EnableVideoOutput(int displayMode, int flags)
        {
            var fn = Marshal.GetDelegateForFunctionPointer<EnableVideoOutputDel>((IntPtr)_vtable[6]);
            return fn(_ptr, displayMode, flags);
        }

        public int DisableVideoOutput()
        {
            var fn = Marshal.GetDelegateForFunctionPointer<DisableVideoOutputDel>((IntPtr)_vtable[7]);
            return fn(_ptr);
        }

        public int CreateAndScheduleFrame(int width, int height, int rowBytes,
            IntPtr pixelData, int pixelDataLen,
            long displayTime, long duration, long timeScale)
        {
            // CreateVideoFrame -> vtable[9]
            var createFn = Marshal.GetDelegateForFunctionPointer<CreateVideoFrameDel>((IntPtr)_vtable[9]);
            int hr = createFn(_ptr, width, height, rowBytes,
                0x32767975, // bmdFormat8BitYUV = '2vuy' = 0x32767975
                0,          // bmdFrameFlagDefault
                out IntPtr framePtr);
            if (hr != 0) return hr;

            // Get the frame vtable to find GetBytes
            // IDeckLinkMutableVideoFrame vtable:
            // 0:QI 1:AddRef 2:Release (IUnknown)
            // 3:GetWidth 4:GetHeight 5:GetRowBytes 6:GetPixelFormat 7:GetFlags
            // 8:GetTimecode 9:GetAncillaryData
            // 10:SetFlags 11:SetTimecode 12:SetTimecodeFromComponents
            // 13:SetAncillaryData 14:SetTimecodeUserBits 15:SetInterfaceProvider
            // But GetBytes is on IDeckLinkVideoBuffer - need to QI for it
            // IDeckLinkVideoBuffer GUID: CCB4B64A-5C86-4E02-B778-885D352709FE
            Guid bufGuid = new Guid("CCB4B64A-5C86-4E02-B778-885D352709FE");
            int qiHr = Marshal.QueryInterface(framePtr, ref bufGuid, out IntPtr bufPtr);
            if (qiHr == 0 && bufPtr != IntPtr.Zero)
            {
                void** framevt = *(void***)bufPtr;
                // IDeckLinkVideoBuffer vtable: 0:QI 1:AddRef 2:Release 3:GetBytes 4:StartAccess 5:EndAccess
                var getBytesFn = Marshal.GetDelegateForFunctionPointer<GetBytesDel>((IntPtr)framevt[3]);
                getBytesFn(bufPtr, out IntPtr dstPtr);
                System.Buffer.MemoryCopy((void*)pixelData, (void*)dstPtr, pixelDataLen, pixelDataLen);
                var relBuf = Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)framevt[2]);
                relBuf(bufPtr);
            }

            // ScheduleVideoFrame -> vtable[12]
            var schedFn = Marshal.GetDelegateForFunctionPointer<ScheduleVideoFrameDel>((IntPtr)_vtable[12]);
            hr = schedFn(_ptr, framePtr, displayTime, duration, timeScale);

            // Release the frame
            void** fvt = *(void***)framePtr;
            var relFrame = Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)fvt[2]);
            relFrame(framePtr);

            return hr;
        }

        public uint GetBufferedVideoFrameCount()
        {
            var fn = Marshal.GetDelegateForFunctionPointer<GetBufferedVideoFrameCountDel>((IntPtr)_vtable[13]);
            fn(_ptr, out uint count);
            return count;
        }

        public int StartScheduledPlayback(long startTime, long timeScale, double speed)
        {
            var fn = Marshal.GetDelegateForFunctionPointer<StartScheduledPlaybackDel>((IntPtr)_vtable[14]);
            return fn(_ptr, startTime, timeScale, speed);
        }

        public int StopScheduledPlayback(long stopTime, long timeScale)
        {
            var fn = Marshal.GetDelegateForFunctionPointer<StopScheduledPlaybackDel>((IntPtr)_vtable[15]);
            return fn(_ptr, stopTime, out _, timeScale);
        }

        public void Dispose()
        {
            if (_ptr != IntPtr.Zero)
            {
                var relFn = Marshal.GetDelegateForFunctionPointer<ReleaseDel>((IntPtr)_vtable[2]);
                relFn(_ptr);
                _ptr = IntPtr.Zero;
            }
        }
    }
}
