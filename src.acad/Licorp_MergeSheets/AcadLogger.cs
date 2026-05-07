using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

namespace Licorp_MergeSheets
{
    public static class AcadLogger
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern void OutputDebugString(string lpOutputString);

        private static readonly string LogFilePath;
        private static readonly object LockObj = new object();

        static AcadLogger()
        {
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Licorp_MergeSheets", "Logs");

            try
            {
                Directory.CreateDirectory(logDir);
                LogFilePath = Path.Combine(logDir, $"MergeLog_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            }
            catch
            {
                LogFilePath = Path.Combine(Path.GetTempPath(), $"MergeLog_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            }
        }

        public static void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var logLine = $"[{timestamp}] {message}";

            WriteToEditor(logLine);
            WriteToDebugView(logLine);
            WriteToFile(logLine);
        }

        public static void LogInfo(string message) => Log($"[INFO] {message}");
        public static void LogError(string message) => Log($"[ERROR] {message}");
        public static void LogDebug(string message) => Log($"[DEBUG] {message}");
        public static void LogWarning(string message) => Log($"[WARN] {message}");

        public static void LogSection(string title)
        {
            var separator = new string('=', 60);
            Log(separator);
            Log($"[SECTION] {title}");
            Log(separator);
        }

        private static void WriteToEditor(string message)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    var ed = doc.Editor;
                    ed.WriteMessage($"\n{message}");
                }
            }
            catch
            {
                // Silently ignore editor errors
            }
        }

        private static void WriteToFile(string message)
        {
            lock (LockObj)
            {
                try
                {
                    File.AppendAllText(LogFilePath, message + Environment.NewLine);
                }
                catch
                {
                    // Silently ignore file errors
                }
            }
        }

        private static void WriteToDebugView(string message)
        {
            try
            {
                // DebugView can capture this from acad.exe/accoreconsole.exe when Win32 capture is enabled.
                OutputDebugString(message);
            }
            catch
            {
                // Silently ignore debug output errors
            }
        }

        public static string GetLogFilePath() => LogFilePath;
    }
}