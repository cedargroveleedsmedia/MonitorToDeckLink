using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using SharpDX.Direct3D11;
using SharpDX.DXGI;

namespace MonitorToDeckLink
{
    // Restored missing classes to fix CS0246 errors
    public class MonitorInfo
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public string DeviceName { get; set; } = "";
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsPrimary { get; set; }
        public override string ToString() => $"{(IsPrimary ? "★ " : "")}{Name} ({Width}x{Height})";
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
        public long TsScale { get; set; }
        public long TsDuration { get; set; }
        public override string ToString() => Label;
    }

    public class AppSettings
    {
        public string LastMonitor { get; set; } = "";
        public string LastDeckLink { get; set; } = "";
        public string LastFormat { get; set; } = "";
    }

    public partial class MainWindow : Window
    {
        private CancellationTokenSource? _cts;
        private Task? _captureTask;
        private readonly string _logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MonitorToDeckLink.log");
        private readonly string _settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");

        private readonly List<OutputFormatInfo> _formats = new()
        {
            new() { Label="1080p 60", ModeInt=0x48703630, Width=1920, Height=1080, TsScale=60000, TsDuration=1000 },
            new() { Label="1080p 59.94", ModeInt=0x48703539, Width=1920, Height=1080, TsScale=60000, TsDuration=1001 },
            new() { Label="1080p 30", ModeInt=0x48703330, Width=1920, Height=1080, TsScale=30000, TsDuration=1000 },
            new() { Label="1080i 60", ModeInt=0x48693630, Width=1920, Height=1080, TsScale=30000, TsDuration=1000 },
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Log("=== App Started ===");
            PopulateMonitors();
            PopulateDeckLinks();
            cmbFormats.ItemsSource = _formats;
            LoadSettings();
        }

        private void Log(string msg)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
            try { File.AppendAllText(_logPath, line + Environment.NewLine); } 
            catch { Debug.WriteLine("Logging failed"); }
            Dispatcher.Invoke(() => { txtLog.Text += line + "\n"; logScroller.ScrollToEnd(); });
        }

        private void LoadSettings()
        {
            if (!File.Exists(_settingsPath)) return;
            try {
                var json = File.ReadAllText(_settingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                if (settings == null) return;

                var m = ((List<MonitorInfo>)cmbMonitors.ItemsSource).FirstOrDefault(x => x.DeviceName == settings.LastMonitor);
                if (m != null) cmbMonitors.SelectedItem = m;

                var d = ((List<DeckLinkDeviceInfo>)cmbDeckLinks.ItemsSource).FirstOrDefault(x => x.Name == settings.LastDeckLink);
                if (d != null) cmbDeckLinks.SelectedItem = d;

                var f = _formats.FirstOrDefault(x => x.Label == settings.LastFormat);
                if (f != null) cmbFormats.SelectedItem = f;
            } catch { }
        }

        private void SaveSettings()
        {
            var settings = new AppSettings {
                LastMonitor = (cmbMonitors.SelectedItem as MonitorInfo)?.DeviceName ?? "",
                LastDeckLink = (cmbDeckLinks.SelectedItem as DeckLinkDeviceInfo)?.Name ?? "",
                LastFormat = (cmbFormats.SelectedItem as OutputFormatInfo)?.Label ?? ""
            };
            try { File.WriteAllText(_settingsPath, JsonSerializer.Serialize(settings)); } catch { }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (cmbMonitors.SelectedItem is not MonitorInfo monitor || cmbDeckLinks.SelectedItem is not DeckLinkDeviceInfo dlInfo || cmbFormats.SelectedItem is not OutputFormatInfo format) return;

            SaveSettings();
            btnStart.IsEnabled = false; btnStop.IsEnabled = true;
            _cts = new CancellationTokenSource();
            int deviceIndex = cmbDeckLinks.SelectedIndex;

            _captureTask = Task.Run(() => {
                var staThread = new Thread(() => {
                    try {
                        var iter = (IDeckLinkIterator2)new CDeckLinkIterator2();
                        IDeckLink2? target = null;
                        for (int i = 0; i <= deviceIndex; i++) iter.Next(out target);
                        if (target != null) {
                            IntPtr iunk = Marshal.GetIUnknownForObject(target);
                            Guid outputGuid = new Guid("1A8077F1-9FE2-4533-8147-2294305E253F");
                            if (Marshal.QueryInterface(iunk, ref outputGuid, out IntPtr outPtr) == 0) {
                                CaptureLoop(monitor, outPtr, format, _cts.Token);
                                Marshal.Release(outPtr);
                            }
                            Marshal.Release(iunk);
                        }
                    } catch (Exception ex) { Log($"Error: {ex.Message}"); }
                });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start(); staThread.Join();
            });
        }

        private void btnStop_Click(object sender, RoutedEventArgs e) => StopCapture();

        private async void StopCapture()
        {
            _cts?.Cancel();
            Log("Stopping output...");
            if (_captureTask != null) await _captureTask;
            Dispatcher.Invoke(() => { btnStart.IsEnabled = true; btnStop.IsEnabled = false; });
            Log("Output stopped.");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => _cts?.Cancel();
        private void btnClearLog_Click(object sender, RoutedEventArgs e) => txtLog.Text = "";
        private void btnCopyLog_Click(object sender, RoutedEventArgs e) { try { Clipboard.SetText(txtLog.Text); } catch { } }
        private void cmbMonitors_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
        private void cmbDeckLinks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }

        private void PopulateMonitors()
        {
            var monitors = new List<MonitorInfo>();
            using var factory = new Factory1();
            int i = 1;
            foreach (var adapter in factory.Adapters1) {
                foreach (var output in adapter.Outputs) {
                    monitors.Add(new MonitorInfo { Name = $"Monitor {i++}", DeviceName = output.Description.DeviceName, Width = output.Description.DesktopBounds.Right, Height = output.Description.DesktopBounds.Bottom });
                    output.Dispose();
                }
                adapter.Dispose();
            }
            cmbMonitors.ItemsSource = monitors;
            cmbMonitors.SelectedIndex = 0;
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
            cmbDeckLinks.SelectedIndex = 0;
        }

        private unsafe void CaptureLoop(MonitorInfo monitor, IntPtr outputRawPtr, OutputFormatInfo format, CancellationToken ct)
        {
            using var deckOutput = new DeckLinkOutput(outputRawPtr) { Logger = msg => Log(msg) };
            using var d3dDevice = new SharpDX.Direct3D11.Device(SharpDX.Direct3D.DriverType.Hardware, DeviceCreationFlags.BgraSupport);
            using var factory = new Factory1();
            
            Output1? dupeOutput = null;
            foreach (var adapter in factory.Adapters1) {
                foreach (var output in adapter.Outputs) {
                    if (output.Description.DeviceName == monitor.DeviceName) { dupeOutput = output.QueryInterface<Output1>(); break; }
                    output.Dispose();
                }
                if (dupeOutput != null) break;
                adapter.Dispose();
            }

            if (dupeOutput == null) return;
            using var deskDupe = dupeOutput.DuplicateOutput(d3dDevice);
            using var stagingTex = new Texture2D(d3dDevice, new Texture2DDescription { Width = format.Width, Height = format.Height, MipLevels = 1, ArraySize = 1, Format = Format.B8G8R8A8_UNorm, SampleDescription = new SampleDescription(1, 0), Usage = ResourceUsage.Staging, CpuAccessFlags = CpuAccessFlags.Read });

            deckOutput.EnableVideoOutput(format.ModeInt, 0);
            var callback = new FrameCallback();
            IntPtr cbPtr = Marshal.GetComInterfaceForObject<FrameCallback, IDeckLinkVideoOutputCallback>(callback);
            deckOutput.SetFrameCallback(cbPtr);

            const int POOL_SIZE = 3;
            IntPtr[] frames = new IntPtr[POOL_SIZE];
            IntPtr[] frameBufs = new IntPtr[POOL_SIZE];

            for (int i = 0; i < POOL_SIZE; i++) {
                if (deckOutput.CreateVideoFrame(format.Width, format.Height, format.Width * 2, 0x32767579, 0, out frames[i]) == 0)
                    deckOutput.GetFrameBytes(frames[i], out frameBufs[i], msg => Log(msg));
            }

            long frameCount = 0;
            while (!ct.IsCancellationRequested) {
                if (deskDupe.TryAcquireNextFrame(10, out _, out SharpDX.DXGI.Resource res).Success) {
                    using (var tex = res.QueryInterface<Texture2D>()) d3dDevice.ImmediateContext.CopyResource(tex, stagingTex);
                    res.Dispose(); deskDupe.ReleaseFrame();

                    var mapped = d3dDevice.ImmediateContext.MapSubresource(stagingTex, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                    int slot = (int)(frameCount % POOL_SIZE);
                    
                    if (frameBufs[slot] != IntPtr.Zero)
                        BgraToUyvy((byte*)mapped.DataPointer, mapped.RowPitch, format.Width, format.Height, (byte*)frameBufs[slot], format.Width, format.Height);
                    
                    d3dDevice.ImmediateContext.UnmapSubresource(stagingTex, 0);
                    deckOutput.ScheduleVideoFrame(frames[slot], frameCount * format.TsDuration, format.TsDuration, format.TsScale);
                    if (frameCount == 2) deckOutput.StartScheduledPlayback(0, format.TsScale, 1.0);
                    frameCount++;
                }
            }
            foreach (var f in frames) deckOutput.ReleaseFrame(f);
        }

        private static unsafe void BgraToUyvy(byte* src, int srcPitch, int srcW, int srcH, byte* dst, int dstW, int dstH)
        {
            if (dst == null) return;
            for (int y = 0; y < dstH; y++) {
                byte* sRow = src + (y * srcPitch);
                byte* dRow = dst + (y * dstW * 2);
                for (int x = 0; x < dstW; x += 2) {
                    byte* p0 = sRow + (x * 4); byte* p1 = sRow + ((x + 1) * 4);
                    dRow[0] = 128; dRow[1] = p0[1]; dRow[2] = 128; dRow[3] = p1[1];
                    dRow += 4;
                }
            }
        }
    }
}
