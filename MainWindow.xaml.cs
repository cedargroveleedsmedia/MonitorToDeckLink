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
        public long TsScale { get; set; }
        public long TsDuration { get; set; }
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
            new() { Label="1080p 23.98", ModeInt=0x32337073, Width=1920, Height=1080, FrameRate=24000.0/1001, TsScale=24000, TsDuration=1001 },
            new() { Label="1080p 24",    ModeInt=0x32347073, Width=1920, Height=1080, FrameRate=24,            TsScale=24,    TsDuration=1 },
            new() { Label="1080p 25",    ModeInt=0x48703235, Width=1920, Height=1080, FrameRate=25,            TsScale=25,    TsDuration=1 },
            new() { Label="1080p 29.97", ModeInt=0x48703239, Width=1920, Height=1080, FrameRate=30000.0/1001, TsScale=30000, TsDuration=1001 },
            new() { Label="1080p 30",    ModeInt=0x48703330, Width=1920, Height=1080, FrameRate=30,            TsScale=30,    TsDuration=1 },
            new() { Label="1080i 50",    ModeInt=0x48693530, Width=1920, Height=1080, FrameRate=25,            TsScale=25,    TsDuration=1 },
            new() { Label="1080i 59.94", ModeInt=0x48693539, Width=1920, Height=1080, FrameRate=30000.0/1001, TsScale=30000, TsDuration=1001 },
            new() { Label="1080i 60",    ModeInt=0x48693630, Width=1920, Height=1080, FrameRate=30,            TsScale=30,    TsDuration=1 },
            new() { Label="1080p 50",    ModeInt=0x48703530, Width=1920, Height=1080, FrameRate=50,            TsScale=50,    TsDuration=1 },
            new() { Label="1080p 59.94", ModeInt=0x48703539, Width=1920, Height=1080, FrameRate=60000.0/1001, TsScale=60000, TsDuration=1001 },
            new() { Label="1080p 60",    ModeInt=0x48703630, Width=1920, Height=1080, FrameRate=60,            TsScale=60,    TsDuration=1 },
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Log($"=== MonitorToDeckLink started {DateTime.Now} ===");
            PopulateMonitors();
            PopulateDeckLinks();
            PopulateFormats();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => StopCapture();

        private void btnClearLog_Click(object sender, RoutedEventArgs e) => txtLog.Text = "";
        private void btnCopyLog_Click(object sender, RoutedEventArgs e) { try { Clipboard.SetText(txtLog.Text); } catch { } }

        // Restored selection handlers required by XAML
        private void cmbMonitors_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
        private void cmbDeckLinks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }

        private void Log(string msg)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
            try { File.AppendAllText(_logPath, line + Environment.NewLine); } catch { }
            Dispatcher.Invoke(() => { txtLog.Text += line + "\n"; logScroller.ScrollToEnd(); });
        }

        private void PopulateMonitors()
        {
            var monitors = new List<MonitorInfo>();
            using var factory = new Factory1();
            foreach (var adapter in factory.Adapters1) {
                foreach (var dxgiOut in adapter.Outputs) {
                    var desc = dxgiOut.Description;
                    monitors.Add(new MonitorInfo {
                        Index = monitors.Count, Name = $"Monitor {monitors.Count + 1}",
                        DeviceName = desc.DeviceName,
                        Width  = desc.DesktopBounds.Right - desc.DesktopBounds.Left,
                        Height = desc.DesktopBounds.Bottom - desc.DesktopBounds.Top,
                        IsPrimary = monitors.Count == 0
                    });
                    dxgiOut.Dispose();
                }
                adapter.Dispose();
            }
            cmbMonitors.ItemsSource = monitors;
            if (monitors.Count > 0) cmbMonitors.SelectedIndex = 0;
        }

        private void PopulateDeckLinks()
        {
            var devices = new List<DeckLinkDeviceInfo>();
            var iterator = (IDeckLinkIterator2)new CDeckLinkIterator2();
            while (iterator.Next(out IDeckLink2 device) == 0 && device != null) {
                device.GetDisplayName(out string name);
                devices.Add(new DeckLinkDeviceInfo { Name = name, Device = device });
            }
            cmbDeckLinks.ItemsSource = devices;
            if (devices.Count > 0) cmbDeckLinks.SelectedIndex = 0;
        }

        private void PopulateFormats() { cmbFormats.ItemsSource = _formats; cmbFormats.SelectedIndex = 10; }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMonitors.SelectedItem is not MonitorInfo monitor) return;
            if (cmbDeckLinks.SelectedItem is not DeckLinkDeviceInfo dlInfo) return;
            if (cmbFormats.SelectedItem is not OutputFormatInfo format) return;

            btnStart.IsEnabled = false; btnStop.IsEnabled = true;
            Log($"--- Starting capture ---");
            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            int deviceIndex = cmbDeckLinks.SelectedIndex;

            var staThread = new Thread(() => {
                try {
                    var iter2 = (IDeckLinkIterator2)new CDeckLinkIterator2();
                    int idx = 0;
                    while (iter2.Next(out IDeckLink2 dev) == 0 && dev != null) {
                        if (idx == deviceIndex) {
                            IntPtr iunk = Marshal.GetIUnknownForObject(dev);
                            Guid preferredGuid = new Guid("1A8077F1-9FE2-4533-8147-2294305E253F");
                            if (Marshal.QueryInterface(iunk, ref preferredGuid, out IntPtr outPtr) == 0) {
                                CaptureLoop(monitor, outPtr, format, token);
                                Marshal.Release(outPtr);
                            }
                            Marshal.Release(iunk);
                            break;
                        }
                        idx++;
                    }
                }
                catch (Exception ex) { Log($"ERROR: {ex.Message}"); }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.IsBackground = true;
            staThread.Start();
            _captureTask = Task.CompletedTask; // Reference variable to fix CS0169
        }

        private void btnStop_Click(object sender, RoutedEventArgs e) => StopCapture();
        private void StopCapture() { _cts?.Cancel(); }

        private unsafe void CaptureLoop(MonitorInfo monitor, IntPtr outputRawPtr, OutputFormatInfo format, CancellationToken ct)
        {
            using var deckOutput = new DeckLinkOutput(outputRawPtr) { Logger = msg => Log(msg) };
            using var d3dDevice = new SharpDX.Direct3D11.Device(SharpDX.Direct3D.DriverType.Hardware, DeviceCreationFlags.BgraSupport);
            using var factory = new Factory1();
            
            Output1? dupeOutput = null;
            foreach (var adpt in factory.Adapters1) {
                foreach (var dxgiOut in adpt.Outputs) {
                    if (dxgiOut.Description.DeviceName == monitor.DeviceName) { dupeOutput = dxgiOut.QueryInterface<Output1>(); break; }
                    dxgiOut.Dispose();
                }
                if (dupeOutput != null) break;
                adpt.Dispose();
            }

            if (dupeOutput == null) throw new Exception("DXGI output not found.");
            using var deskDupe = dupeOutput.DuplicateOutput(d3dDevice);
            using var stagingTex = new Texture2D(d3dDevice, new Texture2DDescription {
                Width = monitor.Width, Height = monitor.Height,
                MipLevels = 1, ArraySize = 1, Format = Format.B8G8R8A8_UNorm,
                SampleDescription = new SampleDescription(1, 0), Usage = ResourceUsage.Staging, CpuAccessFlags = CpuAccessFlags.Read
            });

            deckOutput.EnableVideoOutput(format.ModeInt, 0);
            deckOutput.EnableAudioOutput();

            var callback = new FrameCallback();
            IntPtr callbackPtr = Marshal.GetComInterfaceForObject<FrameCallback, IDeckLinkVideoOutputCallback>(callback);
            deckOutput.SetFrameCallback(callbackPtr);

            const int POOL_SIZE = 3;
            int rowBytes = format.Width * 2;
            var frames = new IntPtr[POOL_SIZE];
            var frameBufs = new IntPtr[POOL_SIZE];

            for (int i = 0; i < POOL_SIZE; i++) {
                // Fixed: Correct FourCC is 0x32767579 ('2vuy')
                int hr = deckOutput.CreateVideoFrame(format.Width, format.Height, rowBytes, 0x32767579, 0, out frames[i]);
                if (hr != 0) throw new Exception($"CreateFrame failed: 0x{hr:X8}");
                deckOutput.GetFrameBytes(frames[i], out frameBufs[i], msg => Log(msg));
            }

            long frameNumber = 0;
            while (!ct.IsCancellationRequested) {
                if (deskDupe.TryAcquireNextFrame(10, out _, out SharpDX.DXGI.Resource res).Success) {
                    using (var t = res.QueryInterface<Texture2D>()) d3dDevice.ImmediateContext.CopyResource(t, stagingTex);
                    deskDupe.ReleaseFrame();
                    res.Dispose();

                    var mapped = d3dDevice.ImmediateContext.MapSubresource(stagingTex, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                    int slot = (int)(frameNumber % POOL_SIZE);
                    BgraToUyvy((byte*)mapped.DataPointer, mapped.RowPitch, monitor.Width, monitor.Height, (byte*)frameBufs[slot], format.Width, format.Height);
                    d3dDevice.ImmediateContext.UnmapSubresource(stagingTex, 0);

                    deckOutput.ScheduleVideoFrame(frames[slot], frameNumber * format.TsDuration, format.TsDuration, format.TsScale);
                    if (frameNumber == 2) deckOutput.StartScheduledPlayback(0, format.TsScale, 1.0);
                    frameNumber++;
                }
            }
            deckOutput.StopScheduledPlayback(0, format.TsScale);
            foreach (var f in frames) deckOutput.ReleaseFrame(f);
        }

        private static unsafe void BgraToUyvy(byte* src, int srcPitch, int srcW, int srcH, byte* dst, int dstW, int dstH)
        {
            float sx = (float)srcW / dstW, sy = (float)srcH / dstH;
            for (int y = 0; y < dstH; y++) {
                byte* sr = src + Math.Min((int)(y * sy), srcH - 1) * srcPitch;
                byte* dr = dst + y * dstW * 2;
                for (int x = 0; x < dstW; x += 2) {
                    byte* p0 = sr + Math.Min((int)(x * sx), srcW - 1) * 4;
                    byte* p1 = sr + Math.Min((int)((x + 1) * sx), srcW - 1) * 4;
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
