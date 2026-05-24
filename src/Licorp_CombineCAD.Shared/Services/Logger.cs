using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.UI;

namespace Licorp_CombineCAD.Services
{
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    public static class Logger
    {
        private static bool _initialized;
        private static string _logPath;
        private static readonly object _lock = new object();

        public static void Initialize()
        {
            if (_initialized) return;

            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Licorp_CombineCAD", "Logs");

            try
            {
                Directory.CreateDirectory(logDir);
            }
            catch
            {
                logDir = Path.GetTempPath();
            }

            _logPath = Path.Combine(logDir, $"ExportLog_{DateTime.Now:yyyyMMdd}.log");
            _initialized = true;

            WriteRaw("INFO", $"Logger initialized. Log directory: {logDir}");
        }

        private static void WriteRaw(string level, string message)
        {
            if (!_initialized || string.IsNullOrEmpty(_logPath)) return;

            try
            {
                var line = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
                lock (_lock)
                {
                    File.AppendAllText(_logPath, line);
                }
            }
            catch
            {
                // Silently fail - logging should never crash the app
            }
        }

        public static void Write(LogLevel level, string message, bool toTaskDialog = false)
        {
            var levelStr = level switch
            {
                LogLevel.Debug => "DBG",
                LogLevel.Info => "INF",
                LogLevel.Warning => "WRN",
                LogLevel.Error => "ERR",
                _ => "INF"
            };

            WriteRaw(levelStr, message);

            if (toTaskDialog || level == LogLevel.Error)
            {
                ShowTaskDialog(level, message);
            }
        }

        public static void LogDebug(string message) => Write(LogLevel.Debug, message);
        public static void LogInfo(string message) => Write(LogLevel.Info, message);
        public static void LogWarning(string message) => Write(LogLevel.Warning, message);
        public static void LogError(string message) => Write(LogLevel.Error, message, true);

        public static void LogSection(string title)
        {
            var separator = new string('=', 60);
            WriteRaw("INF", separator);
            WriteRaw("INF", $"[SECTION] {title}");
            WriteRaw("INF", separator);
        }

        public static void LogProperties(string title, object obj)
        {
            WriteRaw("INF", $"--- {title} ---");
            var type = obj.GetType();
            var props = type.GetProperties();
            foreach (var prop in props)
            {
                try
                {
                    var value = prop.GetValue(obj)?.ToString() ?? "null";
                    WriteRaw("DBG", $"  {prop.Name}: {value}");
                }
                catch
                {
                    WriteRaw("DBG", $"  {prop.Name}: [error reading]");
                }
            }
        }

        public static void LogException(Exception ex, string context = "")
        {
            WriteRaw("ERR", $"[EXCEPTION] {context}: {ex.Message}");
            if (ex.StackTrace != null)
                WriteRaw("ERR", ex.StackTrace);
        }

        public static string GetLogFilePath()
        {
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Licorp_CombineCAD", "Logs");

            if (Directory.Exists(logDir))
            {
                var latestLog = Directory.GetFiles(logDir, "ExportLog_*.log")
                    .OrderByDescending(f => f)
                    .FirstOrDefault();

                if (!string.IsNullOrEmpty(latestLog))
                    return latestLog;
            }

            return "Not initialized";
        }

        public static string GetBufferedLog()
        {
            var logPath = GetLogFilePath();
            if (logPath != "Not initialized" && File.Exists(logPath))
            {
                try
                {
                    return File.ReadAllText(logPath);
                }
                catch { }
            }
            return "No log entries found.";
        }

        public static void ShowBufferedLog(UIApplication uiApp)
        {
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Licorp_CombineCAD", "Logs");

            if (Directory.Exists(logDir))
            {
                var latestLog = Directory.GetFiles(logDir, "ExportLog_*.log")
                    .OrderByDescending(f => f)
                    .FirstOrDefault();

                if (!string.IsNullOrEmpty(latestLog) && File.Exists(latestLog))
                {
                    var content = File.ReadAllText(latestLog);
                    TaskDialog.Show("Export Log", content.Length > 4000 ? content.Substring(0, 4000) + "\n... (truncated)" : content);
                    return;
                }
            }

            TaskDialog.Show("Export Log", "No log entries found.");
        }

        private static void ShowTaskDialog(LogLevel level, string message)
        {
            try
            {
                var td = new TaskDialog("Licorp CombineCAD")
                {
                    TitleAutoPrefix = true,
                    AllowCancellation = true
                };

                switch (level)
                {
                    case LogLevel.Error:
                        td.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                        td.CommonButtons = TaskDialogCommonButtons.Ok;
                        break;
                    case LogLevel.Warning:
                        td.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                        td.CommonButtons = TaskDialogCommonButtons.Ok;
                        break;
                    default:
                        td.MainIcon = TaskDialogIcon.TaskDialogIconNone;
                        td.CommonButtons = TaskDialogCommonButtons.Ok;
                        break;
                }

                td.MainInstruction = $"{level}: {message}";
                td.Show();
            }
            catch { }
        }
    }
}
