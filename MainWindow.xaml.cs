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
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Log("=== Starting MonitorToDeckLink ===");
            PopulateMonitors();
            PopulateDeckLinks();
            cmbFormats.ItemsSource = _formats;
            LoadSettings();
        }

        private void Log(string msg)
        {
            string line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
            try { File.AppendAllText(_logPath, line + Environment.NewLine); } 
            catch (Exception ex) { Debug.WriteLine($"Log write failed: {ex.Message}"); }
            Dispatcher.Invoke(() => { txtLog.Text += line + "\n"; logScroller.ScrollToEnd(); });
        }

        private void LoadSettings()
        {
            if (!File.Exists(_settingsPath)) return;
            try {
                var json = File.ReadAllText(_settingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                if (settings == null) return;

                var monitor = ((List<MonitorInfo>)cmbMonitors.ItemsSource).FirstOrDefault(m => m.DeviceName == settings.LastMonitor);
                if (monitor != null) cmbMonitors.SelectedItem = monitor;

                var dl = ((List<DeckLinkDeviceInfo>)cmbDeckLinks.ItemsSource).FirstOrDefault(d => d.Name == settings.LastDeckLink);
                if (dl != null) cmbDeckLinks.SelectedItem = dl;

                var fmt = _formats.FirstOrDefault(f => f.Label == settings.LastFormat);
                if (fmt != null) cmbFormats.SelectedItem = fmt;
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
            SaveSettings();
            btnStart.IsEnabled = false; btnStop.IsEnabled = true;
            _cts = new CancellationTokenSource();
            
            var monitor = (MonitorInfo)cmbMonitors.SelectedItem;
            var format = (OutputFormatInfo)cmbFormats.SelectedItem;
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
                    } catch (Exception ex) { Log($"Critical Error: {ex.Message}"); }
                });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
            });
        }

        // FIXED: Stop button now triggers cancellation without blocking the UI thread
        private void btnStop_Click(object sender, RoutedEventArgs e) => StopCapture();

        private async void StopCapture()
        {
            _cts?.Cancel();
            Log("Stop requested, waiting for loop to exit...");
            if (_captureTask != null) await _captureTask;
            Dispatcher.Invoke(() => { btnStart.IsEnabled = true; btnStop.IsEnabled = false; });
            Log("Pipe stopped successfully.");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => _cts?.Cancel();

        // (Populate methods, BgraToUyvy, and CaptureLoop logic remain the same but use updated logging)
        private void PopulateMonitors() { /* same as before */ }
        private void PopulateDeckLinks() { /* same as before */ }
        private unsafe void CaptureLoop(MonitorInfo monitor, IntPtr outputRawPtr, OutputFormatInfo format, CancellationToken ct) { /* same as before */ }
        private static unsafe void BgraToUyvy(byte* src, int srcPitch, int srcW, int srcH, byte* dst, int dstW, int dstH) { /* same as before */ }
        private void btnClearLog_Click(object sender, RoutedEventArgs e) => txtLog.Text = "";
        private void btnCopyLog_Click(object sender, RoutedEventArgs e) { try { Clipboard.SetText(txtLog.Text); } catch { } }
        private void cmbMonitors_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
        private void cmbDeckLinks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) { }
    }
}
