using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Merges multiple DWG files into a single DWG with multiple layouts (MultiLayout),
    /// single layout (SingleLayout), or Model Space export.
    /// </summary>
    public class DwgMergeService
    {
        private readonly string _accoreconsolePath;
        private readonly string _acadPath;
        private readonly string _pluginPath;
        private readonly bool _allowFullAutoCadFallback;

        private string _verticalAlign = "Top";
        private string _dwgVersion = "Current";
        private readonly MergeReliabilityOptions _reliability = new MergeReliabilityOptions();
        private int _expectedSheetCount;
        private bool _lastRunReturnedPluginStatus;

        public DwgMergeService(
            string accoreconsolePath = null,
            string acadPath = null)
        {
            _accoreconsolePath = accoreconsolePath ?? AutoCadLocatorService.FindAcCoreConsole();
            _acadPath = acadPath ?? AutoCadLocatorService.FindAutoCAD();

            var pluginDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                @"Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle");

            var enginePath = !string.IsNullOrEmpty(_accoreconsolePath) ? _accoreconsolePath : _acadPath;
            _pluginPath = Path.Combine(pluginDir, "Contents", GetPluginSubFolder(enginePath), "Licorp_MergeSheets.dll");

            // Full AutoCAD fallback can pop UI and is less deterministic with script completion.
            // Default to disabled unless explicitly enabled via env var.
            _allowFullAutoCadFallback = string.Equals(
                Environment.GetEnvironmentVariable("LICORP_ALLOW_FULL_AUTOCAD_FALLBACK"),
                "1",
                StringComparison.Ordinal);
        }

        public bool IsAvailable => !string.IsNullOrEmpty(_accoreconsolePath) || !string.IsNullOrEmpty(_acadPath);

        public bool IsPluginLoaded => File.Exists(_pluginPath);

        public string LastError { get; private set; }

        public string LastLogPath { get; private set; }

        public void SetDwgVersion(string version)
        {
            _dwgVersion = version ?? "Current";
        }

        public void SetVerticalAlignment(string alignment)
        {
            _verticalAlign = alignment ?? "Top";
        }

        public void SetExpectedSheetCount(int expectedSheetCount)
        {
            _expectedSheetCount = Math.Max(0, expectedSheetCount);
        }

        public void EnsurePluginInstalled()
        {
            if (IsPluginLoaded) return;

            try
            {
                var pluginDir = Path.GetDirectoryName(_pluginPath);
                if (!Directory.Exists(pluginDir))
                    Directory.CreateDirectory(pluginDir);

                var sourceDll = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "acad", "Release", "Licorp_MergeSheets.dll");
                if (File.Exists(sourceDll))
                {
                    File.Copy(sourceDll, _pluginPath, true);
                    Trace.WriteLine($"[Merge] Plugin installed to: {_pluginPath}");
                }
                else
                {
                    Trace.WriteLine($"[Merge] Plugin source DLL not found: {sourceDll}");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] Plugin install failed: {ex.Message}");
            }
        }

        public async Task<bool> MergeToMultiLayoutAsync(
            List<string> dwgFiles,
            string outputPath,
            List<string> layoutNames,
            IProgress<MergeProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            // Route MultiLayout through ModelFirstMultiLayout mode in plugin (Mlabs-style behavior)
            return await MergeAsync(dwgFiles, outputPath, layoutNames, "ModelFirstMultiLayout", null, progress, cancellationToken);
        }

        public async Task<bool> MergeToSingleLayoutAsync(
            List<string> dwgFiles,
            string outputPath,
            string layoutName = "Combined",
            IProgress<MergeProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            var layoutNames = dwgFiles == null
                ? null
                : Enumerable.Repeat(layoutName, dwgFiles.Count).ToList();

            return await MergeAsync(dwgFiles, outputPath, layoutNames, "SingleLayout", null, progress, cancellationToken);
        }

        public async Task<bool> MergeToModelSpaceAsync(
            List<string> dwgFiles,
            string outputPath,
            List<string> layoutNames = null,
            IProgress<MergeProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            string seedPath = null;
            try
            {
                var firstValid = dwgFiles?.FirstOrDefault(File.Exists);
                if (string.IsNullOrWhiteSpace(firstValid))
                {
                    LastError = "No valid source files for ModelSpace merge.";
                    return false;
                }

                seedPath = CreateConsoleSeedFile(firstValid);
                return await MergeAsync(dwgFiles, outputPath, layoutNames, "ModelSpace", seedPath, progress, cancellationToken);
            }
            finally
            {
                TryDeleteTempFile(seedPath);
            }
        }

        /// <summary>
        /// Fix DWG files that may have empty ModelSpace (e.g., sheets with only schedules).
        /// Adds a tiny invisible point to ModelSpace to prevent CloneLayoutFromSource errors.
        /// </summary>
        private void FixEmptyModelSpaceFiles(List<string> dwgFiles, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(_accoreconsolePath))
                return;

            foreach (var file in dwgFiles)
            {
                if (!File.Exists(file))
                    continue;

                try
                {
                    var fi = new FileInfo(file);
                    
                    // Always try to fix - schedule-only DWG files can be any size
                    Trace.WriteLine($"[Merge] Fixing empty ModelSpace: {Path.GetFileName(file)} ({fi.Length} bytes)");

                    var scriptPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_FixEmpty_{Guid.NewGuid():N}.scr");
                    // Script to switch to ModelSpace, add a point, save
                    var lines = new List<string>
                    {
                        "TILEMODE",
                        "1",
                        "POINT",
                        "0,0,0",
                        "ZOOM",
                        "EXTENTS",
                        "QSAVE",
                        "QUIT",
                        "Y"
                    };
                    File.WriteAllLines(scriptPath, lines);

                    var startInfo = new ProcessStartInfo
                    {
                        FileName = _accoreconsolePath,
                        Arguments = $"/i \"{file}\" /s \"{scriptPath}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = false,
                        RedirectStandardError = false
                    };

                    using (var process = Process.Start(startInfo))
                    {
                        if (process != null)
                        {
                            const int fixTimeoutMs = 30000;
                            if (!process.WaitForExit(fixTimeoutMs))
                            {
                                try { process.Kill(); }
                                catch { }

                                Trace.WriteLine(
                                    $"[Merge] FixEmptyModelSpace timed out after {fixTimeoutMs / 1000}s for {Path.GetFileName(file)}");
                                continue;
                            }
                        }
                    }

                    TryDeleteTempFile(scriptPath);
                    Trace.WriteLine($"[Merge] Fixed empty ModelSpace: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[Merge] FixEmptyModelSpace failed for {file}: {ex.Message}");
                }
            }
        }

        private async Task<bool> MergeAsync(
            List<string> dwgFiles,
            string outputPath,
            List<string> layoutNames,
            string mode,
            string inputOverride,
            IProgress<MergeProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            LastError = null;
            LastLogPath = null;

            if (dwgFiles == null || dwgFiles.Count == 0)
            {
                LastError = "No source DWG files were supplied.";
                return false;
            }

            if (!IsAvailable)
            {
                LastError = "AutoCAD/AcCoreConsole is not available.";
                Trace.WriteLine("[Merge] AutoCAD engine not available");
                return false;
            }

            var validFiles = dwgFiles.Where(f => !string.IsNullOrWhiteSpace(f) && File.Exists(f)).ToList();
            if (validFiles.Count == 0)
            {
                LastError = "No valid source DWG files exist on disk.";
                Trace.WriteLine("[Merge] No valid source files");
                return false;
            }

            Trace.WriteLine("[Merge] Skipping pre-merge AcCoreConsole ModelSpace fix; layout merger handles empty/schedule sheets internally.");

            string configPath = null;
            string scriptPath = null;
            string statusPath = null;

            try
            {
                EnsurePluginInstalled();

                configPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_{mode}_{Guid.NewGuid():N}.json");
                scriptPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_{mode}_{Guid.NewGuid():N}.scr");
                statusPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_{mode}_{Guid.NewGuid():N}.status.json");

                File.WriteAllText(configPath, CreateMergeConfig(validFiles, layoutNames, outputPath, mode, statusPath));
                CreateMergeScript(scriptPath, configPath);

                Trace.WriteLine($"[ACAD-RUN] mode={mode}");
                Trace.WriteLine($"[ACAD-RUN] configPath={configPath}");
                Trace.WriteLine($"[ACAD-RUN] scriptPath={scriptPath}");
                Trace.WriteLine($"[ACAD-RUN] statusPath={statusPath}");
                Trace.WriteLine($"[ACAD-RUN] outputPath={outputPath}");
                Trace.WriteLine($"[ACAD-RUN] validSources={validFiles.Count}, expectedSheets={_expectedSheetCount}");

                progress?.Report(new MergeProgressInfo
                {
                    Phase = "Merging",
                    CurrentItem = $"Starting {mode} merge...",
                    Current = 0,
                    Total = validFiles.Count
                });

                var inputPath = string.IsNullOrWhiteSpace(inputOverride) ? validFiles[0] : inputOverride;
                var success = await RunMergeEngineAsync(scriptPath, inputPath, outputPath, statusPath, 300000, cancellationToken);
                LastLogPath = LastLogPath ?? GetLatestMergeLogPath();
                return success;
            }
            catch (OperationCanceledException)
            {
                LastError = "Merge cancelled.";
                Trace.WriteLine("[Merge] Cancelled");
                return false;
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                Trace.WriteLine($"[Merge] {mode} error: {ex}");
                return false;
            }
            finally
            {
                TryDeleteTempFile(configPath);
                TryDeleteTempFile(scriptPath);
                TryDeleteTempFile(statusPath);
            }
        }

        private string CreateMergeConfig(List<string> dwgFiles, List<string> layoutNames, string outputPath, string mode, string statusPath)
        {
            var sheetSetIndexPath = Path.Combine(
                Path.GetDirectoryName(outputPath) ?? "",
                Path.GetFileNameWithoutExtension(outputPath) + "_SheetSetIndex.json");

            var config = new MergeConfigDto
            {
                Mode = mode,
                OutputPath = outputPath,
                VerticalAlign = _verticalAlign,
                DwgVersion = _dwgVersion,
                ExpectedSheetCount = _expectedSheetCount > 0 ? _expectedSheetCount : dwgFiles.Count,
                VerifyAfterSave = _reliability.VerifyAfterSave,
                SheetSetEnabled = _reliability.SheetSetEnabled,
                SheetSetIndexPath = sheetSetIndexPath,
                RasterImageMode = _reliability.RasterImageMode,
                StatusPath = statusPath,
                SourceFiles = new List<SourceFileDto>()
            };

            for (int i = 0; i < dwgFiles.Count; i++)
            {
                config.SourceFiles.Add(new SourceFileDto
                {
                    Path = dwgFiles[i],
                    Layout = layoutNames != null && i < layoutNames.Count && !string.IsNullOrWhiteSpace(layoutNames[i])
                        ? layoutNames[i]
                        : $"Layout{i + 1}"
                });
            }

            return JsonConvert.SerializeObject(config, Formatting.Indented);
        }

        private void CreateMergeScript(string scriptPath, string configPath)
        {
            var silentConfigPath = Path.Combine(Path.GetTempPath(), "Licorp_MergeSheets_Config.json");
            File.Copy(configPath, silentConfigPath, true);

            var lines = new List<string>
            {
                "_SECURELOAD",
                "0",
                "NETLOAD",
                $"\"{_pluginPath}\"",
                "_LICORP_MERGESHEETS",
                "QUIT",
                "Y"
            };

            File.WriteAllLines(scriptPath, lines);
            Trace.WriteLine($"[ACAD-RUN] silentConfigPath={silentConfigPath}");
            Trace.WriteLine($"[ACAD-RUN] scriptLines={string.Join(" | ", lines)}");
        }

        private string CreateConsoleSeedFile(string sourcePath)
        {
            var seedPath = Path.Combine(Path.GetTempPath(), $"LicorpCAD_ModelSeed_{Guid.NewGuid():N}.dwg");
            CopyFileShared(sourcePath, seedPath);
            Trace.WriteLine($"[Merge] ModelSpace seed DWG: {seedPath}");
            return seedPath;
        }

        private void CopyFileShared(string sourcePath, string destPath)
        {
            using (var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var dest = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.Read))
            {
                source.CopyTo(dest);
            }
        }

        private async Task<bool> RunMergeEngineAsync(
            string scriptPath,
            string inputPath,
            string outputPath,
            string statusPath,
            int timeoutMs,
            CancellationToken cancellationToken)
        {
            if (!string.IsNullOrWhiteSpace(_accoreconsolePath))
            {
                _lastRunReturnedPluginStatus = false;
                var coreSuccess = await RunAcCoreConsoleInternalAsync(scriptPath, inputPath, outputPath, statusPath, timeoutMs, cancellationToken);
                if (coreSuccess)
                    return true;

                if (_lastRunReturnedPluginStatus)
                {
                    Trace.WriteLine("[Merge] AcCoreConsole plugin completed and returned failure status; not retrying in Full AutoCAD.");
                    return false;
                }

                if (!_allowFullAutoCadFallback)
                {
                    Trace.WriteLine("[Merge] AcCoreConsole failed; Full AutoCAD fallback is disabled.");
                    return false;
                }

                Trace.WriteLine("[Merge] AcCoreConsole failed; trying Full AutoCAD fallback when available.");
            }

            if (!string.IsNullOrWhiteSpace(_acadPath))
                return await RunFullAutoCADInternalAsync(scriptPath, inputPath, outputPath, statusPath, timeoutMs * 2, cancellationToken);

            LastError = LastError ?? "AcCoreConsole failed and Full AutoCAD fallback is not available.";
            return false;
        }

        private async Task<bool> RunAcCoreConsoleInternalAsync(
            string scriptPath,
            string inputPath,
            string outputPath,
            string statusPath,
            int timeoutMs,
            CancellationToken cancellationToken)
        {
            if (!File.Exists(inputPath))
            {
                LastError = $"Input file not found: {inputPath}";
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

            return await RunProcessAndEvaluateAsync("AcCoreConsole", startInfo, outputPath, statusPath, timeoutMs, cancellationToken);
        }

        private async Task<bool> RunFullAutoCADInternalAsync(
            string scriptPath,
            string inputPath,
            string outputPath,
            string statusPath,
            int timeoutMs,
            CancellationToken cancellationToken)
        {
            if (!File.Exists(inputPath))
            {
                LastError = $"Input file not found: {inputPath}";
                return false;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = _acadPath,
                Arguments = $"/nologo \"{inputPath}\" /b \"{scriptPath}\"",
                UseShellExecute = false,
                // Keep Full AutoCAD fallback as unobtrusive as possible when AcCoreConsole is unavailable/fails.
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            return await RunProcessAndEvaluateAsync("Full AutoCAD", startInfo, outputPath, statusPath, timeoutMs, cancellationToken);
        }

        private async Task<bool> RunProcessAndEvaluateAsync(
            string engineName,
            ProcessStartInfo startInfo,
            string outputPath,
            string statusPath,
            int timeoutMs,
            CancellationToken cancellationToken)
        {
            try
            {
                Trace.WriteLine($"[ACAD-RUN] engine={engineName}");
                Trace.WriteLine($"[ACAD-RUN] command={startInfo.FileName} {startInfo.Arguments}");
                Trace.WriteLine($"[ACAD-RUN] timeoutMs={timeoutMs}, statusPath={statusPath}, outputPath={outputPath}");

                using (var process = Process.Start(startInfo))
                {
                    if (process == null)
                    {
                        LastError = $"{engineName} did not start.";
                        return false;
                    }

                    var outputTask = process.StandardOutput.ReadToEndAsync();
                    var errorTask = process.StandardError.ReadToEndAsync();
                    var completed = await Task.Run(() => process.WaitForExit(timeoutMs), cancellationToken);

                    if (!completed)
                    {
                        try { process.Kill(); } catch { }
                        LastError = $"{engineName} timed out after {timeoutMs / 1000} seconds.";
                        return false;
                    }

                    var output = await outputTask;
                    var errors = await errorTask;

                    Trace.WriteLine($"[Merge] {engineName} exit code: {process.ExitCode}");
                    if (!string.IsNullOrWhiteSpace(output))
                        Trace.WriteLine($"[ACAD-RUN] {engineName} stdout: {TrimForTrace(output)}");
                    if (!string.IsNullOrWhiteSpace(errors))
                        Trace.WriteLine($"[ACAD-RUN] {engineName} stderr: {TrimForTrace(errors)}");

                    if (process.ExitCode != 0)
                        Trace.WriteLine($"[ACAD-RUN] {engineName} exited with non-zero code {process.ExitCode}; evaluating status/output before failing.");

                    return await EvaluateMergeRunResultAsync(engineName, process.ExitCode, outputPath, statusPath, cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                LastError = $"{engineName} error: {ex.Message}";
                Trace.WriteLine($"[Merge] {LastError}");
                return false;
            }
        }

        private async Task<bool> EvaluateMergeRunResultAsync(
            string engineName,
            int exitCode,
            string expectedOutputPath,
            string statusPath,
            CancellationToken cancellationToken)
        {
            Trace.WriteLine($"[ACAD-RUN] evaluate-status engine={engineName}, statusPath={statusPath}");

            string deferredFailureMessage = null;

            for (int i = 0; i < 20; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var status = ReadStatus(statusPath);
                if (status != null)
                {
                    _lastRunReturnedPluginStatus = true;
                    LastLogPath = status.LogPath;

                    var finalOutputPath = !string.IsNullOrWhiteSpace(status.OutputPath)
                        ? status.OutputPath
                        : expectedOutputPath;

                    Trace.WriteLine($"[ACAD-RUN] plugin-status success={status.Success}, message={status.Message}, outputPath={status.OutputPath}, logPath={status.LogPath}");

                    if (status.Success)
                    {
                        if (IsLikelyValidCombinedDwg(finalOutputPath, out var reason))
                        {
                            LastError = null;
                            return true;
                        }

                        // Status success can arrive slightly before DWG file flush/handle release.
                        // Defer failure and keep polling for a short period.
                        deferredFailureMessage = $"{engineName} reported success, but output DWG appears invalid: {reason}";
                        Trace.WriteLine($"[ACAD-RUN] status success but output not ready yet (attempt {i + 1}/20): {reason}");
                    }
                    else
                    {
                        LastError = string.IsNullOrWhiteSpace(status.Message)
                            ? $"{engineName} reported failure status."
                            : status.Message;

                        Trace.WriteLine($"[ACAD-RUN] plugin reported failure status, stopping evaluation. lastError={LastError}");
                        return false;
                    }
                }

                if (IsLikelyValidCombinedDwg(expectedOutputPath, out _))
                {
                    LastError = null;
                    Trace.WriteLine($"[ACAD-RUN] status missing/delayed, but output DWG exists and looks valid: {expectedOutputPath}");
                    return true;
                }

                await Task.Delay(500, cancellationToken);
            }

            // Fallback: if plugin log says success and provides output path, trust it when status path mismatches.
            if (TryRecoverSuccessFromLatestLog(expectedOutputPath, out var recoveredOutputPath))
            {
                LastError = null;
                Trace.WriteLine($"[ACAD-RUN] recovered success from merge log; output={recoveredOutputPath}");
                return true;
            }

            if (!string.IsNullOrWhiteSpace(deferredFailureMessage))
            {
                LastError = deferredFailureMessage;
                return false;
            }

            LastLogPath = GetLatestMergeLogPath();
            LastError = $"{engineName} exited with code {exitCode} but no valid status/output was detected.";
            return false;
        }

        private bool TryRecoverSuccessFromLatestLog(string expectedOutputPath, out string recoveredOutputPath)
        {
            recoveredOutputPath = null;

            try
            {
                var logPath = GetLatestMergeLogPath();
                if (string.IsNullOrWhiteSpace(logPath) || !File.Exists(logPath))
                    return false;

                LastLogPath = logPath;
                var lines = File.ReadAllLines(logPath);
                if (lines.Length == 0)
                    return false;

                var hasSuccessLine = lines.Any(l => l.IndexOf("finalStatus success=True", StringComparison.OrdinalIgnoreCase) >= 0)
                    || lines.Any(l => l.IndexOf("Merge Completed Successfully", StringComparison.OrdinalIgnoreCase) >= 0);
                if (!hasSuccessLine)
                    return false;

                var outputLine = lines.LastOrDefault(l => l.IndexOf("Output file:", StringComparison.OrdinalIgnoreCase) >= 0);
                var outputFromLog = outputLine == null
                    ? null
                    : outputLine.Substring(outputLine.IndexOf("Output file:", StringComparison.OrdinalIgnoreCase) + "Output file:".Length).Trim();

                var candidate = !string.IsNullOrWhiteSpace(outputFromLog) ? outputFromLog : expectedOutputPath;
                if (!IsLikelyValidCombinedDwg(candidate, out _))
                    return false;

                recoveredOutputPath = candidate;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private MergeStatusDto ReadStatus(string statusPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(statusPath) || !File.Exists(statusPath))
                    return null;

                var status = JsonConvert.DeserializeObject<MergeStatusDto>(File.ReadAllText(statusPath));
                return status;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Merge] Failed to read status file: {ex.Message}");
                return null;
            }
        }

        private bool IsLikelyValidCombinedDwg(string outputPath, out string reason)
        {
            reason = null;

            try
            {
                if (string.IsNullOrWhiteSpace(outputPath) || !File.Exists(outputPath))
                {
                    reason = "Output file does not exist.";
                    return false;
                }

                var fileInfo = new FileInfo(outputPath);
                if (fileInfo.Length < 4096)
                {
                    reason = $"Output file too small: {fileInfo.Length} bytes.";
                    return false;
                }

                using (var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var header = new byte[6];
                    var read = fs.Read(header, 0, header.Length);
                    if (read < 6)
                    {
                        reason = "Cannot read DWG header.";
                        return false;
                    }

                    var signature = System.Text.Encoding.ASCII.GetString(header);
                    if (!signature.StartsWith("AC", StringComparison.Ordinal))
                    {
                        reason = $"Unexpected DWG signature: {signature}";
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                reason = ex.Message;
                return false;
            }
        }

        private string GetLatestMergeLogPath()
        {
            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Licorp_MergeSheets",
                    "Logs");

                if (!Directory.Exists(logDir))
                    return null;

                return Directory.GetFiles(logDir, "MergeLog_*.log")
                    .Select(p => new FileInfo(p))
                    .OrderByDescending(f => f.LastWriteTimeUtc)
                    .FirstOrDefault()
                    ?.FullName;
            }
            catch
            {
                return null;
            }
        }

        private string GetPluginSubFolder(string enginePath)
        {
            if (!string.IsNullOrEmpty(enginePath))
            {
                var lower = enginePath.ToLowerInvariant();
                if (lower.Contains("2025") || lower.Contains("2026") || lower.Contains("2027"))
                    return "2025";
            }

            return "2024";
        }

        private string TrimForTrace(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            value = value.Trim();
            return value.Length <= 4000 ? value : value.Substring(value.Length - 4000);
        }

        private void TryDeleteTempFile(string path)
        {
            if (string.IsNullOrEmpty(path))
                return;

            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
            }
        }

        private class MergeConfigDto
        {
            public string Mode { get; set; }
            public string OutputPath { get; set; }
            public string VerticalAlign { get; set; }
            public string DwgVersion { get; set; }
            public int ExpectedSheetCount { get; set; }
            public bool VerifyAfterSave { get; set; }
            public bool SheetSetEnabled { get; set; }
            public string SheetSetIndexPath { get; set; }
            public string RasterImageMode { get; set; }
            public string StatusPath { get; set; }
            public List<SourceFileDto> SourceFiles { get; set; }
        }

        private class SourceFileDto
        {
            public string Path { get; set; }
            public string Layout { get; set; }
        }

        private class MergeStatusDto
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public string OutputPath { get; set; }
            public string LogPath { get; set; }
        }

        private class MergeReliabilityOptions
        {
            public bool VerifyAfterSave { get; set; } = true;
            public bool SheetSetEnabled { get; set; } = true;
            public string RasterImageMode { get; set; } = "KeepReference";
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
