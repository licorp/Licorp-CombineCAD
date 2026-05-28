using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Licorp_MergeSheets
{
    /// <summary>
    /// Diagnostics for locked DWG files.
    /// Uses multiple probes:
    /// 1) File metadata and attributes
    /// 2) FileStream open attempts with different FileShare modes
    /// 3) Windows Restart Manager to identify processes that are using the file
    /// </summary>
    internal static class FileLockDiagnostics
    {
        public static string BuildReport(string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("---------------- FILE LOCK DIAGNOSTICS ----------------");
            sb.AppendLine($"Path: {filePath}");
            sb.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            sb.AppendLine($"CurrentProcess: pid={System.Diagnostics.Process.GetCurrentProcess().Id}, name={System.Diagnostics.Process.GetCurrentProcess().ProcessName}");

            try
            {
                sb.AppendLine($"Exists: {File.Exists(filePath)}");
                if (File.Exists(filePath))
                {
                    var fi = new FileInfo(filePath);
                    sb.AppendLine($"Length: {fi.Length} bytes");
                    sb.AppendLine($"LastWriteTime: {fi.LastWriteTime:yyyy-MM-dd HH:mm:ss.fff}");
                    sb.AppendLine($"CreationTime: {fi.CreationTime:yyyy-MM-dd HH:mm:ss.fff}");
                    sb.AppendLine($"Attributes: {fi.Attributes}");
                    sb.AppendLine($"Directory: {fi.DirectoryName}");
                    sb.AppendLine($"DirectoryExists: {Directory.Exists(fi.DirectoryName)}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Metadata error: {ex.GetType().Name}: {ex.Message}");
            }

            ProbeOpen(sb, filePath, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, "Read + Share(ReadWrite|Delete)");
            ProbeOpen(sb, filePath, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete, "ReadWrite + Share(ReadWrite|Delete)");
            ProbeOpen(sb, filePath, FileAccess.ReadWrite, FileShare.None, "ReadWrite + Share(None) / exclusive");

            try
            {
                var lockers = GetLockingProcesses(filePath);
                sb.AppendLine($"RestartManager locking process count: {lockers.Count}");
                foreach (var p in lockers)
                {
                    sb.AppendLine($"Locker: pid={p.ProcessId}, name={p.ApplicationName}, service={p.ServiceShortName}, type={p.ApplicationType}, status={p.AppStatus}, session={p.SessionId}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"RestartManager error: {ex.GetType().Name}: {ex.Message}");
            }

            sb.AppendLine("-------------------------------------------------------");
            return sb.ToString();
        }

        private static void ProbeOpen(StringBuilder sb, string filePath, FileAccess access, FileShare share, string label)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, access, share))
                {
                    sb.AppendLine($"OpenProbe OK: {label}; canRead={fs.CanRead}; canWrite={fs.CanWrite}");
                }
            }
            catch (Exception ex)
            {
                int hresult = Marshal.GetHRForException(ex);
                sb.AppendLine($"OpenProbe FAIL: {label}; {ex.GetType().Name}: {ex.Message}; HResult=0x{hresult:X8}");
            }
        }

        public sealed class LockingProcessInfo
        {
            public int ProcessId { get; set; }
            public string ApplicationName { get; set; }
            public string ServiceShortName { get; set; }
            public RM_APP_TYPE ApplicationType { get; set; }
            public uint AppStatus { get; set; }
            public uint SessionId { get; set; }
        }

        public static List<LockingProcessInfo> GetLockingProcesses(string path)
        {
            uint handle;
            string sessionKey = Guid.NewGuid().ToString("N");
            int result = RmStartSession(out handle, 0, sessionKey);
            if (result != 0)
                throw new Win32Exception(result, "RmStartSession failed");

            try
            {
                string[] resources = { path };
                result = RmRegisterResources(handle, (uint)resources.Length, resources, 0, null, 0, null);
                if (result != 0)
                    throw new Win32Exception(result, "RmRegisterResources failed");

                uint procInfoNeeded = 0;
                uint procInfo = 0;
                uint rebootReasons = 0;

                result = RmGetList(handle, out procInfoNeeded, ref procInfo, null, ref rebootReasons);
                if (result == ERROR_MORE_DATA)
                {
                    var processInfo = new RM_PROCESS_INFO[procInfoNeeded];
                    procInfo = procInfoNeeded;
                    result = RmGetList(handle, out procInfoNeeded, ref procInfo, processInfo, ref rebootReasons);
                    if (result != 0)
                        throw new Win32Exception(result, "RmGetList failed");

                    var list = new List<LockingProcessInfo>();
                    for (int i = 0; i < procInfo; i++)
                    {
                        list.Add(new LockingProcessInfo
                        {
                            ProcessId = processInfo[i].Process.dwProcessId,
                            ApplicationName = processInfo[i].strAppName,
                            ServiceShortName = processInfo[i].strServiceShortName,
                            ApplicationType = processInfo[i].ApplicationType,
                            AppStatus = processInfo[i].AppStatus,
                            SessionId = processInfo[i].SessionId
                        });
                    }
                    return list;
                }

                if (result != 0)
                    throw new Win32Exception(result, "RmGetList failed");

                return new List<LockingProcessInfo>();
            }
            finally
            {
                RmEndSession(handle);
            }
        }

        private const int CCH_RM_MAX_APP_NAME = 255;
        private const int CCH_RM_MAX_SVC_NAME = 63;
        private const int ERROR_MORE_DATA = 234;

        [StructLayout(LayoutKind.Sequential)]
        public struct RM_UNIQUE_PROCESS
        {
            public int dwProcessId;
            public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
        }

        public enum RM_APP_TYPE
        {
            RmUnknownApp = 0,
            RmMainWindow = 1,
            RmOtherWindow = 2,
            RmService = 3,
            RmExplorer = 4,
            RmConsole = 5,
            RmCritical = 1000
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct RM_PROCESS_INFO
        {
            public RM_UNIQUE_PROCESS Process;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
            public string strAppName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
            public string strServiceShortName;
            public RM_APP_TYPE ApplicationType;
            public uint AppStatus;
            public uint TSSessionId;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bRestartable;
            public uint SessionId { get { return TSSessionId; } }
        }

        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        private static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);

        [DllImport("rstrtmgr.dll")]
        private static extern int RmEndSession(uint pSessionHandle);

        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        private static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames, uint nApplications, RM_UNIQUE_PROCESS[] rgApplications, uint nServices, string[] rgsServiceNames);

        [DllImport("rstrtmgr.dll")]
        private static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded, ref uint pnProcInfo, [In, Out] RM_PROCESS_INFO[] rgAffectedApps, ref uint lpdwRebootReasons);
    }
}
