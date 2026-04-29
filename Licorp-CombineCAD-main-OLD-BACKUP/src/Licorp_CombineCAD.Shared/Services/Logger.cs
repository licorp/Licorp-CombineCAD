using System;
using System.IO;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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
        private static readonly object LockObj = new object();
        private static string _logFilePath;
        private static bool _initialized;
        private static StringBuilder _buffer = new StringBuilder();
        private static LogLevel _minLevel = LogLevel.Debug;

        public static void Initialize()
        {
            if (_initialized) return;

            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Licorp_CombineCAD", "Logs");

            try
            {
                Directory.CreateDirectory(logDir);
                _logFilePath = Path.Combine(logDir, $"ExportLog_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                _initialized = true;
                LogInfo($"Logger initialized. Log file: {_logFilePath}");
            }
            catch (Exception ex)
            {
                _logFilePath = Path.Combine(Path.GetTempPath(), $"ExportLog_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                _initialized = true;
                LogError($"Failed to create log directory: {ex.Message}. Using temp path.");
            }
        }

        public static void SetMinLevel(LogLevel level)
        {
            _minLevel = level;
        }

        public static void Log(LogLevel level, string message, bool toTaskDialog = false)
        {
            if (level < _minLevel) return;

            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var levelStr = level.ToString().ToUpper();
            var logLine = $"[{timestamp}] [{levelStr}] {message}";

            WriteToDebug(logLine);
            WriteToFile(logLine);
            WriteToBuffer(logLine);

            if (toTaskDialog || level == LogLevel.Error)
            {
                ShowTaskDialog(level, message);
            }
        }

        public static void LogDebug(string message) => Log(LogLevel.Debug, message);
        public static void LogInfo(string message) => Log(LogLevel.Info, message);
        public static void LogWarning(string message) => Log(LogLevel.Warning, message);
        public static void LogError(string message) => Log(LogLevel.Error, message, true);

        public static void LogSection(string title)
        {
            var separator = new string('=', 60);
            LogInfo(separator);
            LogInfo($"[SECTION] {title}");
            LogInfo(separator);
        }

        public static void LogProperties(string title, object obj)
        {
            LogInfo($"--- {title} ---");
            var type = obj.GetType();
            var props = type.GetProperties();
            foreach (var prop in props)
            {
                try
                {
                    var value = prop.GetValue(obj)?.ToString() ?? "null";
                    LogDebug($"  {prop.Name}: {value}");
                }
                catch
                {
                    LogDebug($"  {prop.Name}: [error reading]");
                }
            }
        }

        public static void LogException(Exception ex, string context = "")
        {
            var sb = new StringBuilder();
            sb.AppendLine($"[EXCEPTION] {context}");
            sb.AppendLine($"Type: {ex.GetType().FullName}");
            sb.AppendLine($"Message: {ex.Message}");

            if (!string.IsNullOrEmpty(ex.StackTrace))
            {
                sb.AppendLine("Stack Trace:");
                sb.AppendLine(ex.StackTrace);
            }

            if (ex.InnerException != null)
            {
                sb.AppendLine($"Inner Exception: {ex.InnerException.Message}");
            }

            LogError(sb.ToString());
        }

        public static string GetLogFilePath() => _logFilePath ?? "Not initialized";

        public static string GetBufferedLog()
        {
            lock (LockObj)
            {
                return _buffer.ToString();
            }
        }

        public static void ShowBufferedLog(UIApplication uiApp)
        {
            var log = GetBufferedLog();
            if (string.IsNullOrEmpty(log))
            {
                TaskDialog.Show("Export Log", "No log entries.");
                return;
            }

            TaskDialog.Show("Export Log", log);
        }

        private static void WriteToDebug(string message)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[LicorpCAD] {message}");
            }
            catch { }
        }

        private static void WriteToFile(string message)
        {
            if (string.IsNullOrEmpty(_logFilePath)) return;

            lock (LockObj)
            {
                try
                {
                    File.AppendAllText(_logFilePath, message + Environment.NewLine);
                }
                catch { }
            }
        }

        private static void WriteToBuffer(string message)
        {
            lock (LockObj)
            {
                _buffer.AppendLine(message);
                const int maxBufferSize = 50000;
                if (_buffer.Length > maxBufferSize)
                {
                    _buffer.Remove(0, _buffer.Length - maxBufferSize);
                }
            }
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