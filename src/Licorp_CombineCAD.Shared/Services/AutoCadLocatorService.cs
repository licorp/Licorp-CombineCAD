using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Win32;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Locates AutoCAD / AcCoreConsole on the system.
    /// Upgraded from Export+ AutoCADBindManager — now finds accoreconsole.exe (headless, faster).
    /// </summary>
    public static class AutoCadLocatorService
    {
        public static string[] SupportedVersions = new[] { "2026", "2025", "2024", "2023", "2022", "2021", "2020" };

        /// <summary>
        /// Get all installed AutoCAD versions
        /// </summary>
        public static List<string> GetInstalledVersions()
        {
            var versions = new List<string>();

            foreach (var ver in SupportedVersions)
            {
                var path = $@"C:\Program Files\Autodesk\AutoCAD {ver}\accoreconsole.exe";
                if (File.Exists(path))
                {
                    versions.Add(ver);
                }
            }

            if (versions.Count == 0)
            {
                var regVersions = GetInstalledVersionsFromRegistry();
                versions.AddRange(regVersions);
            }

            return versions.Distinct().ToList();
        }

        /// <summary>
        /// Find AcCoreConsole.exe (headless AutoCAD engine) — preferred for merge operations
        /// </summary>
        public static string FindAcCoreConsole(string preferredVersion = null)
        {
            // If user specified a version, try that first
            if (!string.IsNullOrEmpty(preferredVersion))
            {
                var specificPath = $@"C:\Program Files\Autodesk\AutoCAD {preferredVersion}\accoreconsole.exe";
                if (File.Exists(specificPath))
                {
                    Debug.WriteLine($"[CombineCAD] Found AcCoreConsole {preferredVersion}: {specificPath}");
                    return specificPath;
                }
            }

            // 1. Try known installation paths (newest first)
            string[] possiblePaths = new[]
            {
                @"C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2023\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2022\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2021\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2020\accoreconsole.exe",
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    Debug.WriteLine($"[CombineCAD] Found AcCoreConsole: {path}");
                    return path;
                }
            }

            // 2. Try Registry lookup
            string regPath = FindFromRegistry("accoreconsole.exe");
            if (!string.IsNullOrEmpty(regPath))
            {
                Debug.WriteLine($"[CombineCAD] Found AcCoreConsole via registry: {regPath}");
                return regPath;
            }

            Debug.WriteLine("[CombineCAD] AcCoreConsole not found on this system");
            return null;
        }

        /// <summary>
        /// Find acad.exe (full AutoCAD) — fallback for bind operations
        /// </summary>
        public static string FindAutoCAD(string preferredVersion = null)
        {
            if (!string.IsNullOrEmpty(preferredVersion))
            {
                var specificPath = $@"C:\Program Files\Autodesk\AutoCAD {preferredVersion}\acad.exe";
                if (File.Exists(specificPath))
                {
                    return specificPath;
                }
            }

            string[] possiblePaths = new[]
            {
                @"C:\Program Files\Autodesk\AutoCAD 2026\acad.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2025\acad.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2024\acad.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2023\acad.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2022\acad.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2021\acad.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2020\acad.exe",
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            return FindFromRegistry("acad.exe");
        }

        /// <summary>
        /// Check if AutoCAD (any version) is available for merge operations
        /// </summary>
        public static bool IsAutoCADAvailable()
        {
            return !string.IsNullOrEmpty(FindAcCoreConsole()) || !string.IsNullOrEmpty(FindAutoCAD());
        }

        /// <summary>
        /// Get AutoCAD version string from the found path (e.g. "2024")
        /// </summary>
        public static string GetAutoCADVersion(string acadPath)
        {
            if (string.IsNullOrEmpty(acadPath)) return null;

            try
            {
                string dirName = Path.GetFileName(Path.GetDirectoryName(acadPath));
                string version = dirName?.Split(' ').LastOrDefault();
                return version;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Get AutoCAD path info for display
        /// </summary>
        public static (bool Available, string Version, string Path) GetAutoCADInfo()
        {
            var accorePath = FindAcCoreConsole();
            if (!string.IsNullOrEmpty(accorePath))
            {
                return (true, GetAutoCADVersion(accorePath), accorePath);
            }

            var acadPath = FindAutoCAD();
            if (!string.IsNullOrEmpty(acadPath))
            {
                return (true, GetAutoCADVersion(acadPath), acadPath);
            }

            return (false, null, null);
        }

        private static string FindFromRegistry(string exeName)
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\AutoCAD"))
                {
                    if (key == null) return null;

                    var subKeys = key.GetSubKeyNames().OrderByDescending(k => k).ToList();

                    foreach (var subKeyName in subKeys)
                    {
                        using (var subKey = key.OpenSubKey(subKeyName))
                        {
                            if (subKey == null) continue;

                            foreach (var installKeyName in subKey.GetSubKeyNames())
                            {
                                using (var installKey = subKey.OpenSubKey(installKeyName))
                                {
                                    var acadLocation = installKey?.GetValue("AcadLocation") as string;
                                    if (!string.IsNullOrEmpty(acadLocation))
                                    {
                                        var exePath = Path.Combine(acadLocation, exeName);
                                        if (File.Exists(exePath))
                                        {
                                            return exePath;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CombineCAD] Registry search failed: {ex.Message}");
            }

            return null;
        }

        private static List<string> GetInstalledVersionsFromRegistry()
        {
            var versions = new List<string>();
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\AutoCAD"))
                {
                    if (key == null) return versions;

                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        var version = subKeyName.Replace("R", "").Split('.').FirstOrDefault();
                        if (!string.IsNullOrEmpty(version) && version.Length == 4)
                        {
                            versions.Add(version);
                        }
                    }
                }
            }
            catch { }
            return versions;
        }
    }
}
