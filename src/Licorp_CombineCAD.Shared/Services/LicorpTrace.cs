// ============================================================
// LicorpTrace.cs — Reusable Debug/Trace Logger for Revit Add-ins
// ============================================================
// USAGE: Copy this single file into any Revit add-in project.
//        Works with RevitAddInManager DockPanel (real-time output).
//        Works in both Debug and Release builds.
//
// REQUIREMENTS:
//   - csproj must define TRACE constant (default in most templates)
//   - No NuGet packages needed
//
// EXAMPLE:
//   LicorpTrace.Info("Export started");
//   LicorpTrace.Warn("File not found");
//   LicorpTrace.Error("Something failed", ex);
//   LicorpTrace.Section("Phase 2: Merge");
// ============================================================

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;

namespace Licorp.Diagnostics
{
    /// <summary>
    /// Lightweight trace logger for Revit add-ins.
    /// Output is captured by RevitAddInManager DockPanel in real-time.
    /// Also writes to a log file for post-mortem debugging.
    /// </summary>
    public static class LicorpTrace
    {
        private static string _logFilePath;
        private static string _prefix = "Licorp";
        private static readonly object _lock = new object();

        // ── Configuration ──────────────────────────────────

        /// <summary>
        /// Initialize with a custom prefix and optional log file.
        /// Call once at startup (e.g., in IExternalApplication.OnStartup).
        /// If not called, defaults are used automatically.
        /// </summary>
        /// <param name="prefix">Prefix shown in output, e.g. "MyCoolAddin"</param>
        /// <param name="enableFileLog">Also write to a log file in %LOCALAPPDATA%</param>
        public static void Init(string prefix = "Licorp", bool enableFileLog = true)
        {
            _prefix = prefix;

            if (enableFileLog)
            {
                try
                {
                    var logDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        prefix, "Logs");
                    Directory.CreateDirectory(logDir);
                    _logFilePath = Path.Combine(logDir, $"Trace_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                }
                catch
                {
                    _logFilePath = Path.Combine(
                        Path.GetTempPath(), $"{prefix}_Trace_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                }
            }
        }

        // ── Public API ─────────────────────────────────────

        /// <summary>General info message</summary>
        public static void Info(string message,
            [CallerMemberName] string caller = "")
        {
            Write("INFO", message, caller);
        }

        /// <summary>Debug-level message (verbose)</summary>
        public static void Dbg(string message,
            [CallerMemberName] string caller = "")
        {
            Write("DBG", message, caller);
        }

        /// <summary>Warning message</summary>
        public static void Warn(string message,
            [CallerMemberName] string caller = "")
        {
            // Prefix "Warning:" triggers yellow color in RevitAddInManager
            Write("Warning", message, caller);
        }

        /// <summary>Error message</summary>
        public static void Error(string message, Exception ex = null,
            [CallerMemberName] string caller = "")
        {
            // Prefix "Error:" triggers red color in RevitAddInManager
            Write("Error", message, caller);
            if (ex != null)
            {
                Write("Error", $"  → {ex.GetType().Name}: {ex.Message}", caller);
                if (ex.InnerException != null)
                    Write("Error", $"  → Inner: {ex.InnerException.Message}", caller);
            }
        }

        /// <summary>Section header for visual grouping</summary>
        public static void Section(string title)
        {
            var sep = new string('═', 50);
            WriteLine($"[{_prefix}] {sep}");
            WriteLine($"[{_prefix}] ▸ {title}");
            WriteLine($"[{_prefix}] {sep}");
        }

        // ── Color keywords for RevitAddInManager ───────────
        // These prefixes trigger colored output in RevitAddInManager DockPanel:
        //   "Warning:"  → Yellow
        //   "Error:"    → Red
        //   "Add:"      → Green
        //   "Modify:"   → Blue
        //   "Delete:"   → Red/Strike

        /// <summary>Green colored output in RevitAddInManager</summary>
        public static void Add(string message) => WriteLine($"Add: [{_prefix}] {message}");

        /// <summary>Blue colored output in RevitAddInManager</summary>
        public static void Modify(string message) => WriteLine($"Modify: [{_prefix}] {message}");

        /// <summary>Red/strike colored output in RevitAddInManager</summary>
        public static void Delete(string message) => WriteLine($"Delete: [{_prefix}] {message}");

        // ── Internal ───────────────────────────────────────

        private static void Write(string level, string message, string caller)
        {
            var time = DateTime.Now.ToString("HH:mm:ss.fff");
            var line = string.IsNullOrEmpty(caller)
                ? $"[{_prefix}] [{time}] [{level}] {message}"
                : $"[{_prefix}] [{time}] [{level}] [{caller}] {message}";
            WriteLine(line);
        }

        private static void WriteLine(string line)
        {
            try
            {
                // Trace.WriteLine works in BOTH Debug and Release builds
                // This is what RevitAddInManager DockPanel captures
                Trace.WriteLine(line);

                // Also write to Debug output (visible in Visual Studio Output window)
                Debug.WriteLine(line);
            }
            catch { }

            // Write to file for post-mortem debugging
            if (!string.IsNullOrEmpty(_logFilePath))
            {
                lock (_lock)
                {
                    try
                    {
                        File.AppendAllText(_logFilePath, line + Environment.NewLine);
                    }
                    catch { }
                }
            }
        }

        /// <summary>Get the current log file path</summary>
        public static string GetLogPath() => _logFilePath ?? "(not initialized)";
    }
}
