using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
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
        public IDeckLink2 Device { get; set; } = null!;
        public override string ToString() => Name;
    }

    public class OutputFormatInfo
    {
        public string Label { get; set; } = "";
        public int ModeInt { get; set; }
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

        // BMDDisplayMode FourCC values
        private readonly List<OutputFormatInfo> _formats = new()
        {
            new() { Label="1080p 23.98", ModeInt=0x48703233, Width=1920, Height=1080, FrameRate=24000.0/1001 },
            new() { Label="1080p 24",    ModeInt=0x48703234, Width=1920, Height=1080, FrameRate=24 },
            new() { Label="1080p 25",    ModeInt=0x48703235, Width=1920, Height=1080, FrameRate=25 },
            new() { Label="1080p 29.97", ModeInt=0x48703239, Width=1920, Height=1080, FrameRate=30000.0/1001 },
            new() { Label="1080p 30",    ModeInt=0x48703330, Width=1920, Height=1080, FrameRate=30 },
            new() { Label="1080i 50",    ModeInt=0x48693530, Width=1920, Height=1080, FrameRate=25 },
            new() { Label="1080i 59.94", ModeInt=0x48693539, Width=1920, Height=1080, FrameRate=30000.0/1001 },
            new() { Label="1080i 60",    ModeInt=0x48693630, Width=1920, Height=1080, FrameRate=30 },
            new() { Label="1080p 50",    ModeInt=0x48703530, Width=1920, Height=1080, FrameRate=50 },
            new() { Label="1080p 59.94", ModeInt=0x48703539, Width=1920, Height=1080, FrameRate=60000.0/1001 },
            new() { Label="1080p 60",    ModeInt=0x48703630, Width=1920, Height=1080, FrameRate=60 },
            new() { Label="720p 50",     ModeInt=0x68703530, Width=1280, Height=720,  FrameRate=50 },
            new() { Label="720p 59.94",  ModeInt=0x68703539, Width=1280, Height=720,  FrameRate=60000.0/1001 },
            new() { Label="720p 60",     ModeInt=0x68703630, Width=1280, Height=720,  FrameRate=60 },
            new() { Label="2160p 23.98", ModeInt=0x34703233, Width=3840, Height=2160, FrameRate=24000.0/1001 },
            new() { Label="2160p 25",    ModeInt=0x34703235, Width=3840, Height=2160, FrameRate=25 },
            new() { Label="2160p 29.97", ModeInt=0x34703239, Width=3840, Height=2160, FrameRate=30000.0/1001 },
            new() { Label="2160p 30",    ModeInt=0x34703330, Width=3840, Height=2160, FrameRate=30 },
        };

        public MainWindow()
        {
            InitializeComponent();
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                Log($"UNHANDLED: {e.ExceptionObject}");
            System.Windows.Application.Current.DispatcherUnhandledException += (s, e) =>
            {
                Log($"DISPATCHER: {e.Exception}");
                e.Handled = true;
            };
            TaskScheduler.UnobservedTaskException += (s, e) =>
            {
                Log($"TASK: {e.Exception}");
                e.SetObserved();
            };
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Log($"=== MonitorToDeckLink started {DateTime.Now} ===");
            Log($"Log: {_logPath}");
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

        private void PopulateMonitors()
        {
            var monitors = new List<MonitorInfo>();
            try
            {
                Log("Enumerating monitors...");
                using var factory = new Factory1();
                int ai = 0;
                foreach (var adapter in factory.Adapters1)
                {
                    int oi = 0;
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
                            IsPrimary = ai == 0 && oi == 0
                        };
                        monitors.Add(m);
                        Log($"  {m.Name}: {desc.DeviceName} {m.Width}x{m.Height}");
                        dxgiOut.Dispose(); oi++;
                    }
                    adapter.Dispose(); ai++;
                }
            }
            catch (Exception ex) { Log($"Monitor error: {ex.Message}"); }
            cmbMonitors.ItemsSource = monitors;
            if (monitors.Count > 0) cmbMonitors.SelectedIndex = 0;
        }

        private void PopulateDeckLinks()
        {
            var devices = new List<DeckLinkDeviceInfo>();
            try
            {
                Log("Enumerating DeckLink devices...");
                var iterator = (IDeckLinkIterator2)new CDeckLinkIterator2();
                while (true)
                {
                    int hr = iterator.Next(out IDeckLink2 device);
                    if (hr != 0 || device == null) break;
                    device.GetDisplayName(out string name);
                    devices.Add(new DeckLinkDeviceInfo { Name = name, Device = device });
                    Log($"  Found: {name}");
                }
                Log($"Found {devices.Count} DeckLink device(s).");
            }
            catch (Exception ex) { Log($"DeckLink error: {ex}"); }
            cmbDeckLinks.ItemsSource = devices;
            if (devices.Count > 0) cmbDeckLinks.SelectedIndex = 0;
        }

        private void PopulateFormats()
        {
            cmbFormats.ItemsSource = _formats;
            cmbFormats.SelectedIndex = 2; // 1080p25
        }

        private void cmbMonitors_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
        private void cmbDeckLinks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMonitors.SelectedItem is not MonitorInfo monitor)        { SetStatus("Select a monitor.", true); return; }
            if (cmbDeckLinks.SelectedItem is not DeckLinkDeviceInfo dlInfo) { SetStatus("Select a DeckLink device.", true); return; }
            if (cmbFormats.SelectedItem is not OutputFormatInfo format)     { SetStatus("Select a format.", true); return; }

            btnStart.IsEnabled = false;
            btnStop.IsEnabled = true;
            SetStatus($"Starting: {monitor.Name} → {dlInfo.Name} @ {format.Label}");
            Log($"--- Starting capture ---");

            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            var device = dlInfo.Device; // capture ref for STA thread

            // DeckLink COM objects must be created and used on the same STA thread
            // We re-enumerate on the STA thread to get a fresh RCW in the right apartment
            int deviceIndex = cmbDeckLinks.SelectedIndex;
            var tcs = new TaskCompletionSource<bool>();
            var staThread = new Thread(() =>
            {
                try
                {
                    // Re-enumerate DeckLink on this STA thread to get proper apartment-local RCW
                    Log("Re-enumerating DeckLink on STA thread...");
                    var iter2 = (IDeckLinkIterator2)new CDeckLinkIterator2();
                    int idx = 0;
                    bool started = false;
                    while (true)
                    {
                        int hr = iter2.Next(out IDeckLink2 dev);
                        if (hr != 0 || dev == null) break;
                        if (idx == deviceIndex)
                        {
                            dev.GetDisplayName(out string n);
                            Log($"STA thread got device: {n}");
                            // Use Marshal.QueryInterface with raw pointer for reliable QI
                            IntPtr iunk = Marshal.GetIUnknownForObject(dev);
                            Guid outputGuid = new Guid("1A8077F1-9FE2-4533-8147-2294305E253F");
                            int qiHr = Marshal.QueryInterface(iunk, ref outputGuid, out IntPtr outPtr);
                            Marshal.Release(iunk);
                            Log($"QI IDeckLinkOutput: hr=0x{qiHr:X8} ptr=0x{outPtr:X}");
                            if (qiHr == 0 && outPtr != IntPtr.Zero)
                            {
                                // Keep outPtr alive - pass directly to CaptureLoop as raw vtable ptr
                                // Do NOT wrap in RCW - that's what causes the crash
                                Log($"Raw IDeckLinkOutput ptr: 0x{outPtr:X}");
                                Log("Starting capture loop...");
                                CaptureLoop(monitor, outPtr, format, token);
                                Marshal.Release(outPtr);
                            }
                            else
                            {
                                throw new Exception($"QI IDeckLinkOutput failed: 0x{qiHr:X8}");
                            }
                            started = true;
                            break;
                        }
                        idx++;
                    }
                    if (!started) throw new Exception("Could not get IDeckLinkOutput on STA thread.");
                    tcs.SetResult(true);
                }
                catch (Exception ex) { tcs.SetException(ex); }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.IsBackground = true;
            staThread.Name = "DeckLinkCapture";
            staThread.Start();
            _captureTask = tcs.Task;

            _captureTask.ContinueWith(t => Dispatcher.Invoke(() =>
            {
                btnStart.IsEnabled = true;
                btnStop.IsEnabled = false;
                if (t.IsFaulted)
                {
                    Log($"ERROR: {t.Exception?.InnerException}");
                    SetStatus($"Error: {t.Exception?.InnerException?.Message}", true);
                }
                else { Log("--- Stopped ---"); SetStatus("Stopped."); }
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

        private unsafe void CaptureLoop(MonitorInfo monitor, IntPtr outputRawPtr,
            OutputFormatInfo format, CancellationToken ct)
        {
            using var deckOutput = new DeckLinkOutput(outputRawPtr);
            Log($"DeckLinkOutput vtable:\n{deckOutput.DumpVtable()}");
            Log("Creating D3D11 device...");
            using var d3dDevice = new SharpDX.Direct3D11.Device(
                SharpDX.Direct3D.DriverType.Hardware, DeviceCreationFlags.BgraSupport);
            using var dxgiDevice = d3dDevice.QueryInterface<SharpDX.DXGI.Device>();
            using var adapter   = dxgiDevice.GetParent<Adapter>();
            using var factory   = adapter.GetParent<Factory1>();

            Output1? dupeOutput = null;
            foreach (var adpt in factory.Adapters1)
            {
                foreach (var dxgiOut in adpt.Outputs)
                {
                    if (dxgiOut.Description.DeviceName == monitor.DeviceName)
                    { dupeOutput = dxgiOut.QueryInterface<Output1>(); dxgiOut.Dispose(); goto found; }
                    dxgiOut.Dispose();
                }
                adpt.Dispose();
            }
            found:
            if (dupeOutput == null) throw new Exception($"DXGI output not found: {monitor.DeviceName}");
            using var deskDupe = dupeOutput.DuplicateOutput(d3dDevice);
            dupeOutput.Dispose();

            using var stagingTex = new Texture2D(d3dDevice, new Texture2DDescription
            {
                Width = monitor.Width, Height = monitor.Height,
                MipLevels = 1, ArraySize = 1, Format = Format.B8G8R8A8_UNorm,
                SampleDescription = new SampleDescription(1, 0),
                Usage = ResourceUsage.Staging, BindFlags = BindFlags.None,
                CpuAccessFlags = CpuAccessFlags.Read, OptionFlags = ResourceOptionFlags.None
            });

            Log($"Enabling video output: {format.Label} mode=0x{format.ModeInt:X8}");
            int enableHr = deckOutput.EnableVideoOutput(format.ModeInt, 0);
            Log($"EnableVideoOutput returned: 0x{enableHr:X8}");
            if (enableHr != 0) throw new Exception($"EnableVideoOutput failed: 0x{enableHr:X8}");
            Log("Video output enabled!");

            long   frameDuration = (long)(TimeSpan.TicksPerSecond / format.FrameRate);
            double framePeriodMs = 1000.0 / format.FrameRate;
            var    sw            = Stopwatch.StartNew();
            long   frameNumber   = 0;
            byte[]? lastFrame    = null;
            int    timeouts      = 0;

            // Pre-buffer frames then start playback
            long tsScale = (long)Math.Round(format.FrameRate * 1000);
            long tsDur   = 1000;

            Log("Entering capture loop...");
            Log("Waiting for first frame from DXGI...");
            while (!ct.IsCancellationRequested)
            {
                double targetMs = frameNumber * framePeriodMs;
                double wait = targetMs - sw.Elapsed.TotalMilliseconds;
                if (wait > 1) Thread.Sleep((int)wait - 1);
                while (sw.Elapsed.TotalMilliseconds < targetMs) { }

                bool got = false;
                try
                {
                    // Try up to 5 times with 20ms wait to get a frame
                    SharpDX.DXGI.Resource? res = null;
                    for (int attempt = 0; attempt < 5 && res == null; attempt++)
                    {
                        var hr = deskDupe.TryAcquireNextFrame(20,
                            out OutputDuplicateFrameInformation fi,
                            out SharpDX.DXGI.Resource r);
                        if (hr.Success && r != null) { res = r; }
                        else if (attempt == 0 && frameNumber < 3)
                            Log($"TryAcquireNextFrame attempt {attempt}: hr={hr.Code}");
                    }
                    if (res != null)
                    {
                        using (res)
                        using (var t = res.QueryInterface<Texture2D>())
                            d3dDevice.ImmediateContext.CopyResource(t, stagingTex);
                        deskDupe.ReleaseFrame();
                        got = true; timeouts = 0;
                        if (frameNumber == 0) Log("First DXGI frame captured.");
                    }
                    else timeouts++;
                }
                catch (SharpDX.SharpDXException ex)
                {
                    if (timeouts == 0) Log($"TryAcquireNextFrame exception: {ex.ResultCode} {ex.Message}");
                    timeouts++;
                }

                int    rowBytes = format.Width * 2;
                byte[] uyvy     = new byte[rowBytes * format.Height];

                if (got)
                {
                    if (frameNumber == 0) Log("MapSubresource...");
                    var mapped = d3dDevice.ImmediateContext.MapSubresource(
                        stagingTex, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                    if (frameNumber == 0) Log($"Mapped ptr=0x{mapped.DataPointer:X} pitch={mapped.RowPitch}");
                    try
                    {
                        fixed (byte* dst = uyvy)
                            BgraToUyvy((byte*)mapped.DataPointer, mapped.RowPitch,
                                monitor.Width, monitor.Height, dst, format.Width, format.Height);
                    }
                    finally { d3dDevice.ImmediateContext.UnmapSubresource(stagingTex, 0); }
                    if (frameNumber == 0) Log("BgraToUyvy done.");
                    lastFrame = uyvy;
                }
                else if (lastFrame != null) uyvy = lastFrame;
                else { frameNumber++; continue; }

                int createHr = deckOutput.CreateVideoFrame(
                    format.Width, format.Height, rowBytes,
                    0x32767975, 0,
                    out IntPtr framePtr);

                if (frameNumber == 0) Log($"CreateVideoFrame hr=0x{createHr:X8} ptr=0x{framePtr:X}");

                if (createHr == 0 && framePtr != IntPtr.Zero)
                {
                    deckOutput.GetFrameBytes(framePtr, out IntPtr dst, msg => Log(msg));
                    if (dst != IntPtr.Zero)
                    {
                        fixed (byte* src = uyvy)
                            System.Buffer.MemoryCopy(src, (void*)dst, uyvy.Length, uyvy.Length);
                    }
                    deckOutput.EndFrameAccess();
                    int schedHr = deckOutput.ScheduleVideoFrame(framePtr,
                        frameNumber * tsDur, tsDur, tsScale);
                    if (frameNumber < 3) Log($"ScheduleVideoFrame[{frameNumber}] hr=0x{schedHr:X8}");
                    deckOutput.ReleaseFrame(framePtr);
                }
                else if (frameNumber < 3) Log($"CreateVideoFrame hr=0x{createHr:X8}");

                frameNumber++;

                // Start playback after pre-buffering 2 frames
                if (frameNumber == 2)
                {
                    Log("StartScheduledPlayback (after pre-buffering)...");
                    int startHr = deckOutput.StartScheduledPlayback(0, tsScale, 1.0);
                    Log($"StartScheduledPlayback hr=0x{startHr:X8}");
                    if (startHr != 0) Log($"WARNING: StartScheduledPlayback failed: 0x{startHr:X8}");
                }

                if (frameNumber % (long)format.FrameRate == 0)
                {
                    deckOutput.GetBufferedVideoFrameCount(out uint buf);
                    string s = $"Running — frame {frameNumber}  buf:{buf}  timeouts:{timeouts}";
                    Dispatcher.Invoke(() => txtStatus.Text = s);
                    if (frameNumber % ((long)format.FrameRate * 5) == 0) Log(s);
                }
            }

            Log("Stopping...");
            deckOutput.StopScheduledPlayback(0, tsScale);
            deckOutput.DisableVideoOutput();
            Log("Done.");
        }

        private static unsafe void BgraToUyvy(
            byte* src, int srcPitch, int srcW, int srcH,
            byte* dst, int dstW, int dstH)
        {
            float sx = (float)srcW / dstW, sy = (float)srcH / dstH;
            for (int y = 0; y < dstH; y++)
            {
                byte* sr = src + Math.Min((int)(y * sy), srcH - 1) * srcPitch;
                byte* dr = dst + y * dstW * 2;
                for (int x = 0; x < dstW; x += 2)
                {
                    byte* p0 = sr + Math.Min((int)(x * sx),       srcW - 1) * 4;
                    byte* p1 = sr + Math.Min((int)((x+1) * sx),   srcW - 1) * 4;
                    float y0 = 16+0.1826f*p0[2]+0.6142f*p0[1]+0.0620f*p0[0];
                    float y1 = 16+0.1826f*p1[2]+0.6142f*p1[1]+0.0620f*p1[0];
                    float cb = 128-0.1006f*p0[2]-0.3386f*p0[1]+0.4392f*p0[0];
                    float cr = 128+0.4392f*p0[2]-0.3989f*p0[1]-0.0403f*p0[0];
                    dr[0]=(byte)Math.Clamp((int)cb,0,255);
                    dr[1]=(byte)Math.Clamp((int)y0,16,235);
                    dr[2]=(byte)Math.Clamp((int)cr,0,255);
                    dr[3]=(byte)Math.Clamp((int)y1,16,235);
                    dr+=4;
                }
            }
        }
    }
}
