using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// DWG cleanup service — detects and removes XREF files created by Revit export.
    /// Copied and refactored from Export+ DWGCleanupManager.
    /// </summary>
    public static class DwgCleanupService
    {
        /// <summary>
        /// Check if a DWG file has companion XREF files
        /// </summary>
        public static bool HasXRefFiles(string dwgPath)
        {
            if (string.IsNullOrEmpty(dwgPath) || !File.Exists(dwgPath))
                return false;

            try
            {
                var directory = Path.GetDirectoryName(dwgPath);
                var baseName = Path.GetFileNameWithoutExtension(dwgPath);

                var xrefFiles = Directory.GetFiles(directory, "*.dwg")
                    .Where(f =>
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        return !name.Equals(baseName, StringComparison.OrdinalIgnoreCase)
                               && name.StartsWith(baseName, StringComparison.OrdinalIgnoreCase);
                    })
                    .ToList();

                return xrefFiles.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Delete all XREF files associated with a main DWG file
        /// </summary>
        public static int CleanupXRefFiles(string mainDwgPath)
        {
            if (string.IsNullOrEmpty(mainDwgPath) || !File.Exists(mainDwgPath))
                return 0;

            int deletedCount = 0;

            try
            {
                var directory = Path.GetDirectoryName(mainDwgPath);
                var mainFileName = Path.GetFileNameWithoutExtension(mainDwgPath);

                var relatedFiles = Directory.GetFiles(directory, "*.dwg")
                    .Where(f =>
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        return name.StartsWith(mainFileName, StringComparison.OrdinalIgnoreCase);
                    })
                    .OrderBy(f => new FileInfo(f).Length)
                    .ToList();

                if (relatedFiles.Count <= 1)
                    return 0;

                foreach (var file in relatedFiles)
                {
                    if (file.Equals(mainDwgPath, StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        File.Delete(file);
                        deletedCount++;
                        Trace.WriteLine($"[Cleanup] Deleted XREF: {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"[Cleanup] Failed to delete {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Cleanup] Error: {ex.Message}");
            }

            return deletedCount;
        }

        /// <summary>
        /// Cleanup a temporary export folder
        /// </summary>
        public static void CleanupTempFolder(string tempPath)
        {
            if (string.IsNullOrEmpty(tempPath) || !Directory.Exists(tempPath))
                return;

            try
            {
                if (Directory.Exists(tempPath))
                {
                    Directory.Delete(tempPath, true);
                    Trace.WriteLine($"[Cleanup] Deleted temp folder: {tempPath}");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Cleanup] Failed to delete temp folder: {ex.Message}");
            }
        }

        /// <summary>
        /// Get list of XREF file paths for a main DWG
        /// </summary>
        public static string[] GetXRefFiles(string mainDwgPath)
        {
            if (string.IsNullOrEmpty(mainDwgPath) || !File.Exists(mainDwgPath))
                return Array.Empty<string>();

            try
            {
                var directory = Path.GetDirectoryName(mainDwgPath);
                var mainFileName = Path.GetFileNameWithoutExtension(mainDwgPath);

                return Directory.GetFiles(directory, "*.dwg")
                    .Where(f =>
                    {
                        var name = Path.GetFileNameWithoutExtension(f);
                        return !name.Equals(mainFileName, StringComparison.OrdinalIgnoreCase)
                               && name.StartsWith(mainFileName, StringComparison.OrdinalIgnoreCase);
                    })
                    .ToArray();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }
    }
}