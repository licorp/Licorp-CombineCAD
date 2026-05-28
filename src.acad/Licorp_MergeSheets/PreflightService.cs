using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Licorp_MergeSheets
{
    public class PreflightService
    {
        private static readonly string[] DwgSignatures = { "AC10", "AC20" };

        public PreflightResult RunPreflight(MergeConfig config)
        {
            var result = new PreflightResult { Success = true };

            if (config == null)
            {
                result.Success = false;
                result.Errors.Add("Merge config is null");
                return result;
            }

            if (config.SourceFiles == null || config.SourceFiles.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(config.SourceFolder))
                {
                    if (!Directory.Exists(config.SourceFolder))
                    {
                        result.Success = false;
                        result.Errors.Add($"Source folder not found: {config.SourceFolder}");
                        return result;
                    }

                    var folderService = new BatchFolderService();
                    var scannedFiles = folderService.ScanFolder(config.SourceFolder, config.SourcePattern, config.RecursiveScan);
                    config.SourceFiles = scannedFiles;
                    result.TotalFileCount = scannedFiles.Count;
                }
                else
                {
                    result.Success = false;
                    result.Errors.Add("No source files or source folder specified");
                    return result;
                }
            }
            else
            {
                result.TotalFileCount = config.SourceFiles.Count;
            }

            if (string.IsNullOrWhiteSpace(config.OutputPath))
            {
                result.Success = false;
                result.Errors.Add("Output path is not specified");
                return result;
            }

            string outputDir = Path.GetDirectoryName(config.OutputPath);
            if (!string.IsNullOrWhiteSpace(outputDir) && !Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                    result.Warnings.Add($"Created output directory: {outputDir}");
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Errors.Add($"Cannot create output directory: {ex.Message}");
                    return result;
                }
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(outputDir))
                {
                    var root = Path.GetPathRoot(outputDir);
                    if (!string.IsNullOrEmpty(root) && !root.StartsWith("\\\\") && !root.StartsWith("//"))
                    {
                        var drive = new DriveInfo(root);
                        long totalSourceSize = config.SourceFiles?.Where(f => f != null && !string.IsNullOrWhiteSpace(f.Path) && File.Exists(f.Path))
                            .Sum(f => new FileInfo(f.Path).Length) ?? 0;
                        long minRequired = Math.Max(50L * 1024 * 1024, totalSourceSize * 2);

                        if (drive.AvailableFreeSpace < minRequired)
                        {
                            result.Success = false;
                            result.Errors.Add($"Insufficient disk space on {root}. Required: {minRequired / (1024 * 1024)} MB, Available: {drive.AvailableFreeSpace / (1024 * 1024)} MB");
                            return result;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"Could not check disk space: {ex.Message}");
            }

            int validCount = 0;
            var validSources = new List<SourceFile>();

            foreach (var source in config.SourceFiles)
            {
                if (source == null || string.IsNullOrWhiteSpace(source.Path))
                {
                    result.Warnings.Add("Skipping null or empty source file entry");
                    continue;
                }

                if (!File.Exists(source.Path))
                {
                    result.Warnings.Add($"File not found: {source.Path}");
                    continue;
                }

                try
                {
                    var fileInfo = new FileInfo(source.Path);
                    if (fileInfo.Length < 100)
                    {
                        result.Warnings.Add($"File too small ({fileInfo.Length} bytes): {source.Path}");
                        continue;
                    }

                    if (!ValidateDwgHeader(source.Path))
                    {
                        result.Warnings.Add($"File does not appear to be a valid DWG: {source.Path}");
                        continue;
                    }

                    if (IsFileLocked(source.Path))
                    {
                        result.Warnings.Add($"File is locked by another process: {source.Path}");
                        AcadLogger.LogWarning($"Preflight: File is locked by another process: {source.Path}");
                        AcadLogger.LogWarning(FileLockDiagnostics.BuildReport(source.Path));
                        continue;
                    }

                    validSources.Add(source);
                    validCount++;
                }
                catch (Exception ex)
                {
                    result.Warnings.Add($"Error checking file '{source.Path}': {ex.Message}");
                }
            }

            result.ValidFileCount = validCount;
            config.SourceFiles = validSources;

            if (validCount == 0)
            {
                result.Success = false;
                result.Errors.Add("No valid source files found");
                return result;
            }

            if (validCount < result.TotalFileCount)
            {
                result.Warnings.Add($"{result.TotalFileCount - validCount} file(s) skipped (not found or invalid)");
            }

            if (!string.Equals(config.Mode, "MultiLayout", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(config.Mode, "SingleLayout", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(config.Mode, "ModelSpace", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(config.Mode, "ModelFirstMultiLayout", StringComparison.OrdinalIgnoreCase))
            {
                result.Success = false;
                result.Errors.Add($"Unknown merge mode: {config.Mode}");
            }

            if (File.Exists(config.OutputPath))
            {
                result.Warnings.Add($"Output file will be overwritten: {config.OutputPath}");
            }

            AcadLogger.LogInfo($"PreflightService: {result.TotalFileCount} total, {result.ValidFileCount} valid, " +
                $"{result.Errors.Count} error(s), {result.Warnings.Count} warning(s)");

            return result;
        }

        private bool ValidateDwgHeader(string filePath)
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (stream.Length < 6) return false;

                    var header = new byte[6];
                    int bytesRead = stream.Read(header, 0, 6);
                    if (bytesRead < 6) return false;

                    string signature = System.Text.Encoding.ASCII.GetString(header, 0, 4);
                    return DwgSignatures.Any(s => signature.StartsWith(s, StringComparison.OrdinalIgnoreCase));
                }
            }
            catch
            {
                return true;
            }
        }

        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
