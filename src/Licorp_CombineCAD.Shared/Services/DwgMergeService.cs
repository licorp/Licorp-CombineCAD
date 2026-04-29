using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Merges multiple DWG files into a single DWG with multiple layouts (MultiLayout),
    /// single layout (SingleLayout), or Model Space export.
    /// Core feature — requires AutoCAD/AcCoreConsole + Licorp_MergeSheets plugin.
    /// </summary>
public class DwgMergeService
{
    private readonly string _accoreconsolePath;
    private readonly string _pluginPath;
    private string _verticalAlign = "Top";
    private string _dwgVersion = "Current";

public DwgMergeService(string accoreconsolePath = null)
{
_accoreconsolePath = accoreconsolePath ?? AutoCadLocatorService.FindAcCoreConsole();

var pluginDir = Path.Combine(
Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
@"Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle");

string subFolder = "2024";
if (!string.IsNullOrEmpty(_accoreconsolePath))
{
var accoreDir = Path.GetDirectoryName(_accoreconsolePath).ToLowerInvariant();
if (accoreDir.Contains("2025") || accoreDir.Contains("2026") || accoreDir.Contains("2027"))
{
subFolder = "2025";
}
}
_pluginPath = Path.Combine(pluginDir, "Contents", subFolder, "Licorp_MergeSheets.dll");
}

    public bool IsAvailable => !string.IsNullOrEmpty(_accoreconsolePath);

    public bool IsPluginLoaded => File.Exists(_pluginPath);

    public void SetDwgVersion(string version)
    {
        _dwgVersion = version ?? "Current";
    }

    public void SetVerticalAlignment(string alignment)
        {
            _verticalAlign = alignment ?? "Top";
        }

        /// <summary>
        /// Ensure the merge plugin is installed for AcCoreConsole
        /// </summary>
        public void EnsurePluginInstalled()
        {
            if (IsPluginLoaded) return;

            try
            {
                var pluginDir = Path.GetDirectoryName(_pluginPath);
                if (!Directory.Exists(pluginDir))
                {
                    Directory.CreateDirectory(pluginDir);
                }

                var sourceDll = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "acad", "Release", "Licorp_MergeSheets.dll");
                if (File.Exists(sourceDll))
                {
                    File.Copy(sourceDll, _pluginPath, true);
                    Trace.WriteLine($"[Merge] Plugin installed to: {_pluginPath}");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] Plugin install failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Merge multiple DWG files into one file with multiple layouts (each source = 1 layout)
        /// </summary>
        public async Task<bool> MergeToMultiLayoutAsync(
            List<string> dwgFiles,
            string outputPath,
            List<string> layoutNames,
            IProgress<MergeProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (dwgFiles == null || dwgFiles.Count == 0)
                return false;

            if (!IsAvailable)
            {
                Trace.WriteLine("[Merge] AcCoreConsole not available");
                return false;
            }

            var validFiles = dwgFiles.Where(f => File.Exists(f)).ToList();
            if (validFiles.Count == 0)
            {
                Trace.WriteLine("[Merge] No valid source files");
                return false;
            }

            try
            {
                EnsurePluginInstalled();

                var config = CreateMergeConfig(validFiles, layoutNames, outputPath, "MultiLayout");
                var configPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_Merge_{Guid.NewGuid()}.json");
                var scriptPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_Merge_{Guid.NewGuid()}.scr");

                File.WriteAllText(configPath, config);
                CreateMergeScript(scriptPath, configPath, outputPath);

                progress?.Report(new MergeProgressInfo
                {
                    Phase = "Merging",
                    CurrentItem = "Starting merge...",
                    Current = 0,
                    Total = validFiles.Count
                });

                var success = await RunAcCoreConsoleAsync(scriptPath, validFiles[0], outputPath, 300000, cancellationToken);

                try { File.Delete(configPath); } catch { }
                try { File.Delete(scriptPath); } catch { }

                return success;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] Error: {ex.Message}");
                return false;
            }
        }

        public async Task<bool> MergeToSingleLayoutAsync(
            List<string> dwgFiles,
            string outputPath,
            string layoutName = "Combined",
            IProgress<MergeProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (dwgFiles == null || dwgFiles.Count == 0)
                return false;

            if (!IsAvailable)
                return false;

            var validFiles = dwgFiles.Where(f => File.Exists(f)).ToList();
            if (validFiles.Count == 0)
                return false;

            try
            {
                EnsurePluginInstalled();

                var layoutNames = Enumerable.Repeat(layoutName, validFiles.Count).ToList();
                var config = CreateMergeConfig(validFiles, layoutNames, outputPath, "SingleLayout");
                var configPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_Single_{Guid.NewGuid()}.json");
                var scriptPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_Single_{Guid.NewGuid()}.scr");

                File.WriteAllText(configPath, config);
                CreateMergeScript(scriptPath, configPath, outputPath);

                var success = await RunAcCoreConsoleAsync(scriptPath, validFiles[0], outputPath, 300000, cancellationToken);

                try { File.Delete(configPath); } catch { }
                try { File.Delete(scriptPath); } catch { }

                return success;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] SingleLayout error: {ex.Message}");
                return false;
            }
        }

        public async Task<bool> MergeToModelSpaceAsync(
            List<string> dwgFiles,
            string outputPath,
            IProgress<MergeProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (dwgFiles == null || dwgFiles.Count == 0)
                return false;

            if (!IsAvailable)
                return false;

            var validFiles = dwgFiles.Where(f => File.Exists(f)).ToList();
            if (validFiles.Count == 0)
                return false;

            try
            {
                EnsurePluginInstalled();

                var config = CreateMergeConfig(validFiles, null, outputPath, "ModelSpace");
                var configPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_Model_{Guid.NewGuid()}.json");
                var scriptPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_Model_{Guid.NewGuid()}.scr");

                File.WriteAllText(configPath, config);
                CreateMergeScript(scriptPath, configPath, outputPath);

                var success = await RunAcCoreConsoleAsync(scriptPath, validFiles[0], outputPath, 300000, cancellationToken);

                try { File.Delete(configPath); } catch { }
                try { File.Delete(scriptPath); } catch { }

                return success;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] ModelSpace error: {ex.Message}");
                return false;
            }
        }

private string CreateMergeConfig(List<string> dwgFiles, List<string> layoutNames, string outputPath, string mode)
    {
        var sb = new StringBuilder();
        sb.AppendLine("{");
        sb.AppendLine($" \"Mode\": \"{mode}\",");
        sb.AppendLine($" \"OutputPath\": \"{outputPath.Replace("\\", "\\\\")}\",");
        sb.AppendLine($" \"VerticalAlign\": \"{_verticalAlign}\",");
        sb.AppendLine($" \"DwgVersion\": \"{_dwgVersion}\",");
        sb.AppendLine(" \"SourceFiles\": [");

        for (int i = 0; i < dwgFiles.Count; i++)
        {
            var comma = i < dwgFiles.Count - 1 ? "," : "";
            sb.AppendLine($" {{ \"Path\": \"{dwgFiles[i].Replace("\\", "\\\\")}\", \"Layout\": \"{(layoutNames != null && i < layoutNames.Count ? layoutNames[i] : $"Layout{i + 1}")}\" }}{comma}");
        }

sb.AppendLine(" ]");
        sb.AppendLine("}");
        return sb.ToString();
    }

private void CreateMergeScript(string scriptPath, string configPath, string outputPath)
    {
        var sb = new StringBuilder();

        var silentConfigPath = Path.Combine(Path.GetTempPath(), "Licorp_MergeSheets_Config.json");
        File.Copy(configPath, silentConfigPath, true);

        sb.AppendLine("_SECURELOAD");
        sb.AppendLine("0");
        sb.AppendLine("NETLOAD");
        sb.AppendLine($"\"{_pluginPath}\"");
        sb.AppendLine("_LICORP_MERGESHEETS");

        sb.AppendLine("QUIT");
        sb.AppendLine("Y");
        File.WriteAllText(scriptPath, sb.ToString());
    }

        private async Task<bool> RunAcCoreConsoleAsync(string scriptPath, string inputPath, string outputPath, int timeoutMs, CancellationToken cancellationToken)
        {
            try
            {
if (!string.IsNullOrEmpty(_accoreconsolePath))
            {
                return await RunAcCoreConsoleInternalAsync(scriptPath, inputPath, outputPath, timeoutMs, cancellationToken);
            }

            Trace.WriteLine("[Merge] AcCoreConsole not available");
            return false;
            }
            catch (OperationCanceledException)
            {
                Trace.WriteLine("[Merge] Cancelled");
                return false;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] Error: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> RunAcCoreConsoleInternalAsync(string scriptPath, string inputPath, string outputPath, int timeoutMs, CancellationToken cancellationToken)
        {
            try
            {
                if (!File.Exists(inputPath))
                {
                    Trace.WriteLine($"[Merge] Input file not found: {inputPath}");
                    return false;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = _accoreconsolePath,
                    Arguments = $"/i \"{inputPath}\" /s \"{scriptPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(startInfo))
                {
                    if (process == null)
                        return false;

                    var outputTask = process.StandardOutput.ReadToEndAsync();
                    var errorTask = process.StandardError.ReadToEndAsync();

                    var completed = await Task.Run(() => process.WaitForExit(timeoutMs), cancellationToken);

                    if (!completed)
                    {
                        try { process.Kill(); } catch { }
                        return false;
                    }

                    string output = await outputTask;
                    string errors = await errorTask;

                    Trace.WriteLine($"[Merge] AcCoreConsole exit code: {process.ExitCode}");
                    if (!string.IsNullOrEmpty(errors))
                        Trace.WriteLine($"[Merge] Errors: {errors}");

                    if (process.ExitCode == 0 && File.Exists(outputPath))
                        return true;

                    if (!File.Exists(outputPath))
                        Trace.WriteLine($"[Merge] Output file not found: {outputPath}");

                    return false;
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] Console error: {ex.Message}");
                return false;
            }
}
}

public class MergeProgressInfo
    {
        public string Phase { get; set; }
        public string CurrentItem { get; set; }
        public int Current { get; set; }
        public int Total { get; set; }
        public double Percentage => Total > 0 ? (double)Current / Total * 100 : 0;
    }
}
