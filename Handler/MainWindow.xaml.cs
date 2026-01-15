using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Handler
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<ProcessInfo> processList = new();
        private CancellationTokenSource cts = new();
        private bool isRunning = false;
        private long serverTotalRAMBytes = 0;

        // P/Invoke for GUI resources
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetGuiResources(IntPtr hProcess, int uiFlags);

        private const int GR_USER = 0;  // USER objects
        private const int GR_GDI = 1;  // GDI objects

        public MainWindow()
        {
            InitializeComponent();
            dgProcesses.ItemsSource = processList;

            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            CalculateServerTotalRAM();
        }

        private void CalculateServerTotalRAM()
        {
            try
            {
                var searcher = new ManagementObjectSearcher("SELECT Capacity FROM Win32_PhysicalMemory");
                long total = 0;
                foreach (ManagementObject obj in searcher.Get())
                    total += Convert.ToInt64(obj["Capacity"]);

                serverTotalRAMBytes = total;

                Dispatcher.Invoke(() =>
                {
                    double gb = Math.Round(total / (1024.0 * 1024 * 1024), 1);
                    txtServerTotalRAM.Text = $"{gb:N1} GB ({total:N0} bytes)";
                });
            }
            catch (Exception ex)
            {
                txtServerTotalRAM.Text = $"Error: {ex.Message}";
            }
        }

        private async void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            await RefreshOnce();
        }

        private async void chkAutoRefresh_Checked(object sender, RoutedEventArgs e)
        {
            if (isRunning) return;
            isRunning = true;
            cts = new CancellationTokenSource();

            try
            {
                while (!cts.IsCancellationRequested)
                {
                    await RefreshOnce();
                    if (int.TryParse((cmbInterval.SelectedItem as ComboBoxItem)?.Content?.ToString(), out int sec))
                        await Task.Delay(TimeSpan.FromSeconds(sec), cts.Token);
                    else
                        await Task.Delay(5000, cts.Token);
                }
            }
            catch (OperationCanceledException) { }
            finally { isRunning = false; }
        }

        private void chkAutoRefresh_Unchecked(object sender, RoutedEventArgs e)
        {
            cts.Cancel();
            cts = new CancellationTokenSource();
        }

        private async Task RefreshOnce()
        {
            // Show refreshing state
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = "Refreshing...";
                txtStatus.Foreground = Brushes.DarkOrange;
                btnRefresh.IsEnabled = false;
            });

            string filter = txtProcessName.Text?.Trim();

            if (string.IsNullOrWhiteSpace(filter))
            {
                ShowStatus("Enter a process name prefix (e.g. notepad, chrome, videoos)", Brushes.IndianRed);
                Dispatcher.Invoke(() => btnRefresh.IsEnabled = true);
                return;
            }

            try
            {
                // ────────────────────────────────────────────────────────────────
                // 1. Find matching processes (name starts with filter, case-insensitive)
                // ────────────────────────────────────────────────────────────────
                var matchingProcesses = Process.GetProcesses()
                    .Where(p =>
                    {
                        try
                        {
                            return p.ProcessName.StartsWith(filter, StringComparison.OrdinalIgnoreCase);
                        }
                        catch
                        {
                            return false; // skip access-denied/zombie processes
                        }
                    })
                    .Select(CreateProcessInfo)
                    .Where(info => info != null)
                    .OrderByDescending(info => info.HandleCount)
                    .ToList();

                // ────────────────────────────────────────────────────────────────
                // 2. Calculate totals for displayed (filtered) processes
                // ────────────────────────────────────────────────────────────────
                long sumHandles = matchingProcesses.Sum(p => (long)p.HandleCount);
                long sumPaged = matchingProcesses.Sum(p => p.PagedPoolKB);
                long sumPagedPeak = matchingProcesses.Sum(p => p.PagedPoolPeakKB);
                long sumThreads = matchingProcesses.Sum(p => (long)p.ThreadCount);

                // ────────────────────────────────────────────────────────────────
                // 3. Calculate system-wide USER + GDI handles
                // ────────────────────────────────────────────────────────────────
                long totalUserHandles = 0;
                long totalGdiHandles = 0;

                foreach (Process proc in Process.GetProcesses())
                {
                    try
                    {
                        if (proc.Handle != IntPtr.Zero)
                        {
                            totalUserHandles += GetGuiResources(proc.Handle, GR_USER);
                            totalGdiHandles += GetGuiResources(proc.Handle, GR_GDI);
                        }
                    }
                    catch
                    {
                        // Skip access denied / protected processes
                    }
                }

                // Practical max for USER objects (windows, menus, etc.)
                const long PRACTICAL_USER_MAX = 32768L;
                long estimatedRemaining = Math.Max(0, PRACTICAL_USER_MAX - totalUserHandles);

                double usagePercent = totalUserHandles > 0
                    ? Math.Min(100.0, Math.Round((double)totalUserHandles / PRACTICAL_USER_MAX * 100, 1))
                    : 0.0;

                // ────────────────────────────────────────────────────────────────
                // 4. Update UI on dispatcher thread
                // ────────────────────────────────────────────────────────────────
                Dispatcher.Invoke(() =>
                {
                    // Clear and repopulate the list
                    processList.Clear();
                    foreach (var info in matchingProcesses)
                    {
                        processList.Add(info);
                    }

                    // Filtered totals
                    txtTotalHandles.Text = sumHandles.ToString("N0");
                    txtTotalPaged.Text = sumPaged.ToString("N0");
                    txtTotalPagedPeak.Text = sumPagedPeak.ToString("N0");
                    txtTotalThreads.Text = sumThreads.ToString("N0");
                    txtTotalUserHandles.Text = totalUserHandles.ToString("N0");
                    txtUserRemaining.Text = $"{estimatedRemaining:N0} remaining";

                    // System totals
                    txtSystemUserHandles.Text = totalUserHandles.ToString("N0");
                    txtSystemGdiHandles.Text = totalGdiHandles.ToString("N0");

                    // Progress bar & percentage
                    pbUserHandles.Value = usagePercent;
                    txtUserPercent.Text = $"{usagePercent:F1}%";

                    // Color coding
                    SolidColorBrush barColor;
                    if (usagePercent < 70)
                        barColor = new SolidColorBrush(Colors.Green);
                    else if (usagePercent < 90)
                        barColor = new SolidColorBrush(Colors.Orange);
                    else
                        barColor = new SolidColorBrush(Colors.Red);

                    pbUserHandles.Foreground = barColor;
                    txtUserPercent.Foreground = barColor;

                    // Final status
                    txtStatus.Text = $"Found {matchingProcesses.Count} processes starting with '{filter}' • Last update: {DateTime.Now:HH:mm:ss}";
                    txtStatus.Foreground = Brushes.DarkGreen;
                });
            }
            catch (Exception ex)
            {
                ShowStatus($"Error during refresh: {ex.Message}", Brushes.IndianRed);
            }
            finally
            {
                Dispatcher.Invoke(() => btnRefresh.IsEnabled = true);
            }
        }

        private void ShowStatus(string msg, Brush color)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = msg;
                txtStatus.Foreground = color;
            });
        }

        private ProcessInfo? CreateProcessInfo(Process p)
        {
            try
            {
                long paged = 0, pagedPeak = 0;
                try
                {
                    using var counterPaged = new System.Diagnostics.PerformanceCounter("Process", "Pool Paged Bytes", p.ProcessName, true);
                    paged = (long)counterPaged.NextValue();

                    using var counterPeak = new System.Diagnostics.PerformanceCounter("Process", "Pool Paged Bytes Peak", p.ProcessName, true);
                    pagedPeak = (long)counterPeak.NextValue();
                }
                catch { }

                return new ProcessInfo
                {
                    Pid = p.Id,
                    ProcessName = p.ProcessName,
                    HandleCount = p.HandleCount,
                    ThreadCount = p.Threads.Count,
                    PagedPoolKB = paged / 1024,
                    PagedPoolPeakKB = pagedPeak / 1024,
                    StartTime = p.StartTime
                };
            }
            catch
            {
                return null;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            cts.Cancel();
            base.OnClosed(e);
        }
    }

    public class ProcessInfo
    {
        public int Pid { get; set; }
        public string? ProcessName { get; set; }
        public int HandleCount { get; set; }
        public long PagedPoolKB { get; set; }
        public long PagedPoolPeakKB { get; set; }
        public int ThreadCount { get; set; }
        public DateTime StartTime { get; set; }
    }
}