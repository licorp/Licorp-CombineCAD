using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.Services;
using Licorp_CombineCAD.Views;
namespace Licorp_CombineCAD.ViewModels
{
    public partial class ExportDialogViewModel
    {
        private void BrowseFolder()
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select Output Folder",
                SelectedPath = OutputFolder ?? "",
                ShowNewFolderButton = true
            };

            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                OutputFolder = dialog.SelectedPath;
        }

        private bool CanExport()
        {
            return !IsExporting &&
                !IsLoadingSheets &&
                SelectedCount > 0 &&
                !string.IsNullOrEmpty(OutputFolder);
        }

        private void ValidateExportMode()
        {
            StatusMessage = IsAutoCADAvailable
                ? StatusMessage
                : "AutoCAD not detected. Only individual export will be performed.";
        }

        private void CheckAutoCADAvailability()
        {
            var info = AutoCadLocatorService.GetAutoCADInfo();
            IsAutoCADAvailable = info.Available;
            AutoCADVersion = info.Version;
            AutoCADPath = info.Path;

            AvailableAutoCADVersions.Clear();
            foreach (var ver in AutoCadLocatorService.GetInstalledVersions())
                AvailableAutoCADVersions.Add(ver);

            if (!string.IsNullOrEmpty(AutoCADVersion) && !AvailableAutoCADVersions.Contains(AutoCADVersion))
                AvailableAutoCADVersions.Insert(0, AutoCADVersion);

            SelectedAutoCADVersion = AutoCADVersion ?? AvailableAutoCADVersions.FirstOrDefault();

            OnPropertyChanged(nameof(IsAutoCADAvailable));
            OnPropertyChanged(nameof(AutoCADVersion));
            OnPropertyChanged(nameof(AutoCADPath));
            ValidateExportMode();
        }

        private async Task ExecuteExportAsync()
        {
            if (IsExporting)
                return;

            if (!string.IsNullOrEmpty(OutputFolder) && !Directory.Exists(OutputFolder))
            {
                var result = MessageBox.Show(
                    "Output folder does not exist. Create it?",
                    "Folder Not Found", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;

                try { Directory.CreateDirectory(OutputFolder); }
                catch (Exception dirEx)
                {
                    MessageBox.Show("Cannot create folder: " + dirEx.Message, "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            var selectedSheets = AllSheets.Where(s => s.IsSelected).Select(s => s.Model).ToList();
            var settings = BuildExportSettings();
            if (!Directory.Exists(settings.OutputFolder))
                Directory.CreateDirectory(settings.OutputFolder);

            // --- Preflight (ExternalEvent raise #1) ---
            SheetPreflightResult preflightResult;
            try
            {
                preflightResult = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    _sheetCollector.HydrateSheetsForExport(selectedSheets);
                    return _preflightService.Analyze(selectedSheets, settings);
                });
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Preflight] Failed: " + ex);
                MessageBox.Show("Preflight check failed:\n" + ex.Message,
                    "Preflight Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!ConfirmPreflight(preflightResult))
                return;

            IsExporting = true;
            _cancellationTokenSource = new CancellationTokenSource();

            ProgressVm = new ProgressViewModel(() => _cancellationTokenSource?.Cancel());
            ProgressVm.Update("Preparing", "Initializing export...", 0, selectedSheets.Count * 2);
            ProgressVm.StartTimer();

            // Force UI update before handing control to Revit thread.
            await _uiDispatcher.BeginInvoke(new Action(() => { }));

            try
            {
                Trace.WriteLine($"[ExportInit] Begin ExecuteExportAsync | selectedSheets={selectedSheets.Count} | setup='{settings.DwgExportSetupName}' | mode={settings.ExportMode}");

                // BuildExportOptions touches ExportDWGSettings (Revit API) — must run on Revit thread.
                // However, calling it here on the UI thread before the export raise avoids needing
                // a separate ExternalEvent raise, matching the reference implementation pattern.
                // NOTE: BuildExportOptions only reads settings, no transaction needed.
                ProgressVm.Update("Preparing", "Building DWG export options...", 1, selectedSheets.Count * 2);
                var options = await _revitThreadService.RunOnRevitThreadAsync(app =>
                    _exportService.BuildExportOptions(settings));

                if (options == null)
                    throw new InvalidOperationException("Failed to build DWG export options.");

                var cts = _cancellationTokenSource;
                var progressVm = ProgressVm;
                var totalSheets = selectedSheets.Count;

                // --- Export (ExternalEvent raise #2) ---
                ProgressVm.Update("Preparing", "Starting Revit export engine...", 2, selectedSheets.Count * 2);
                ExportResult exportResult = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    var progress = new DirectProgress<ExportProgressInfo>(info =>
                    {
                        _uiDispatcher.BeginInvoke(new Action(() =>
                            progressVm.Update(info.Phase, info.CurrentItem, info.Current, totalSheets * 2)));
                    });

                    return _exportService.ExportSheetsIndividually(
                        selectedSheets, settings, options,
                        progress,
                        cts.Token);
                });
                Trace.WriteLine("[ExportInit] Export complete");

                var exportedFiles = exportResult?.ExportedFiles ?? new List<string>();

                if (cts.Token.IsCancellationRequested)
                {
                    StatusMessage = "Export cancelled.";
                    ProgressVm.StopTimer();
                    ProgressVm.Completed = true;
                    return;
                }

                var selectedCount = selectedSheets.Count;
                var exportedCount = exportedFiles.Count;
                var failedCount = exportResult?.FailedSheets?.Count ?? 0;
                var skippedCount = exportResult?.SkippedSheets?.Count ?? 0;

                if (exportResult != null && (exportResult.HasWarnings || exportedCount < selectedCount))
                {
                    var warningMsg = string.Format(
                        "Selected {0} sheet(s). Exported {1}. Failed {2}. Skipped {3}.",
                        selectedCount,
                        exportedCount,
                        failedCount,
                        skippedCount);

                    if (exportedCount > 0 && exportedCount < selectedCount)
                        warningMsg += "\n\nThe process will continue with successfully exported sheets.";

                    if (exportResult.FailedSheets.Count > 0)
                        warningMsg += "\n\nFailed: " + string.Join(", ", exportResult.FailedSheets);
                    if (exportResult.SkippedSheets.Count > 0)
                        warningMsg += "\n\nSkipped: " + string.Join(", ", exportResult.SkippedSheets);

                    await _uiDispatcher.BeginInvoke(new Action(() =>
                    {
                        MessageBox.Show(warningMsg, "Export Warnings", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }));
                }

                string fileToOpen = null;

                if (exportedFiles.Count > 0)
                {
                    if (IsAutoCADAvailable)
                    {
                        ProgressVm.UpdatePhase("Merging");

                        var accorePath = AutoCadLocatorService.FindAcCoreConsole(SelectedAutoCADVersion);
                        var acadPath = AutoCadLocatorService.FindAutoCAD(SelectedAutoCADVersion);
                        var mergeService = new DwgMergeService(accorePath, acadPath);
                        if (!mergeService.IsAvailable)
                        {
                            StatusMessage = string.Format("Exported {0} individual DWG files. AutoCAD merge engine not available.", exportedFiles.Count);
                            if (OpenAfterExport && exportedFiles.Count > 0)
                                fileToOpen = exportedFiles.First();
                            goto AfterMerge;
                        }

                        mergeService.SetVerticalAlignment(settings.VerticalAlign.ToString());
                        mergeService.SetDwgVersion(settings.DwgVersion ?? "Current");
                        mergeService.SetExpectedSheetCount(exportedFiles.Count);
                        mergeService.SetMergeLayers(settings.MergeLayers);
                        mergeService.SetSheetSortMode(settings.SortMode.ToString());

                        var exportedSheets = exportResult?.ExportedSheets ?? new List<SheetInfo>();
                        // Apply LayoutNameTemplate with {SheetNumber}, {SheetName}, {PaperSize} placeholders
                        var layoutNameTemplate = settings.LayoutNameTemplate ?? "{SheetNumber} - {SheetName}";
                        var layoutNames = exportedSheets.Select(s =>
                            layoutNameTemplate
                                .Replace("{SheetNumber}", s.SheetNumber ?? "")
                                .Replace("{SheetName}", s.SheetName ?? "")
                                .Replace("{PaperSize}", s.PaperSize ?? "")
                        ).ToList();
                        if (layoutNames.Count != exportedFiles.Count)
                            layoutNames = exportedFiles.Select(Path.GetFileNameWithoutExtension).ToList();
                        layoutNames = BuildUniqueLayoutNames(layoutNames);

                        // Collect paper sizes from Revit sheet info for accurate paper plot size in merged DWG
                        var paperSizes = exportedSheets.Select(s => s.PaperSize).ToList();
                        if (paperSizes.Count != exportedFiles.Count)
                            paperSizes = Enumerable.Repeat<string>(null, exportedFiles.Count).ToList();
                        Logger.LogDebug($"[Merge] Paper sizes from Revit: {string.Join(", ", paperSizes.Select(p => p ?? "null"))}");

                        var outputPath = GetUniqueOutputPath();

                        var mergeSuccess = false;

                        var mergeProgress = new DirectProgress<MergeProgressInfo>(info =>
                        {
                            _uiDispatcher.BeginInvoke(new Action(() =>
                            {
                                progressVm.Update(info.Phase, info.CurrentItem, totalSheets + info.Current, totalSheets * 2);
                            }));
                        });

                        switch (ExportMode)
                        {
                            case ExportMode.MultiLayout:
                                mergeSuccess = await mergeService.MergeToMultiLayoutAsync(exportedFiles, outputPath, layoutNames, paperSizes, mergeProgress, cts.Token);
                                break;
                            case ExportMode.SingleLayout:
                                mergeSuccess = await mergeService.MergeToSingleLayoutAsync(exportedFiles, outputPath, "Combined", paperSizes, mergeProgress, cts.Token);
                                break;
                            case ExportMode.ModelSpace:
                                mergeSuccess = await mergeService.MergeToModelSpaceAsync(exportedFiles, outputPath, layoutNames, paperSizes, mergeProgress, cts.Token);
                                break;
                        }

                        if (mergeSuccess)
                        {
                            StatusMessage = exportedCount < selectedCount
                                ? string.Format(
                                    "Partial success: selected {0}, exported/merged {1}. Output: {2}",
                                    selectedCount,
                                    exportedCount,
                                    Path.GetFileName(outputPath))
                                : string.Format("Merged {0} files to {1}", exportedFiles.Count, Path.GetFileName(outputPath));

                            var finalOutputPath = outputPath;
                            if (!DwgMergeService.IsLikelyValidCombinedDwg(finalOutputPath, out var validateReason))
                            {
                                var logPath = mergeService.LastLogPath;
                                await _uiDispatcher.BeginInvoke(new Action(() =>
                                {
                                    MessageBox.Show(
                                        "AutoCAD reported success but the output file still looks invalid.\n\n"
                                        + validateReason
                                        + (string.IsNullOrWhiteSpace(logPath) ? "" : "\n\nLog: " + logPath),
                                        "Merge Warning",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Warning);
                                }));
                            }

                            if (OpenAfterExport)
                                fileToOpen = outputPath;
                        }
                        else
                        {
                            StatusMessage = "Combine failed. Individual DWG files were exported, but the combined DWG is not valid.";
                            var logPath = mergeService.LastLogPath;
                            var detail = string.IsNullOrWhiteSpace(mergeService.LastError)
                                ? "See merge log for details."
                                : mergeService.LastError;

                            await _uiDispatcher.BeginInvoke(new Action(() =>
                            {
                                MessageBox.Show(
                                    detail + (string.IsNullOrWhiteSpace(logPath) ? "" : "\n\nLog: " + logPath),
                                    "Combine Failed",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Warning);
                            }));
                        }
                    }
                    else
                    {
                        StatusMessage = string.Format("Exported {0} individual DWG files. AutoCAD not available for merge.", exportedFiles.Count);
                        if (OpenAfterExport && exportedFiles.Count > 0)
                            fileToOpen = exportedFiles.First();
                    }

                AfterMerge: ;
                }
                else
                {
                    StatusMessage = string.Format(
                        "No sheets were exported successfully (selected {0}, failed {1}, skipped {2}).",
                        selectedCount,
                        failedCount,
                        skippedCount);

                    await _uiDispatcher.BeginInvoke(new Action(() =>
                    {
                        MessageBox.Show(
                            "No sheet was exported successfully, so merge was skipped.\n\n"
                            + string.Format("Selected: {0}\nFailed: {1}\nSkipped: {2}", selectedCount, failedCount, skippedCount),
                            "Export Completed With No Output",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }));
                }

                ProgressVm.StopTimer();
                ProgressVm.Completed = true;
                SaveSettings();
                await Task.Delay(500);

                if (fileToOpen != null)
                    await OpenWithAutoCADAsync(fileToOpen);
            }
            catch (OperationCanceledException)
            {
                StatusMessage = "Export cancelled.";
                Trace.WriteLine("[Export] Cancelled by user");
            }
            catch (TimeoutException tex)
            {
                StatusMessage = "Error: Revit export timed out. Please try again.";
                Trace.WriteLine("[Export] Timeout: " + tex);
                var errorLog = Path.Combine(Path.GetTempPath(), "Licorp_ExportErrors.log");
                File.AppendAllText(errorLog, DateTime.Now.ToString("s") + " Timeout: " + tex + Environment.NewLine);
                await _uiDispatcher.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(
                        "The export operation timed out. This can happen if Revit is busy.\n\nPlease try again or restart Revit.",
                        "Export Timeout",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }));
            }
            catch (Exception ex)
            {
                StatusMessage = "Error: " + ex.Message;
                Trace.WriteLine("[Export] Error: " + ex);
                var errorLog = Path.Combine(Path.GetTempPath(), "Licorp_ExportErrors.log");
                File.AppendAllText(errorLog, DateTime.Now.ToString("s") + " Error: " + ex + Environment.NewLine);
            }
            finally
            {
                IsExporting = false;
                if (ProgressVm != null && !ProgressVm.Completed)
                {
                    ProgressVm.StopTimer();
                    ProgressVm.Completed = true;
                    await Task.Delay(500);
                }

                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                RefreshDerivedState();
            }
        }

        private string GetUniqueOutputPath()
        {
            var baseName = BuildProjectFolderName();
            var path = Path.Combine(GetResolvedOutputFolder(), baseName + ".dwg");

            var counter = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(GetResolvedOutputFolder(), baseName + "_" + counter + ".dwg");
                counter++;
            }

            return path;
        }

        private static List<string> BuildUniqueLayoutNames(List<string> layoutNames)
        {
            var result = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rawName in layoutNames ?? new List<string>())
            {
                var baseName = SanitizeLayoutName(rawName);
                var candidate = baseName;
                var suffix = 1;

                while (!used.Add(candidate))
                {
                    suffix++;
                    candidate = $"{baseName}_{suffix}";
                }

                result.Add(candidate);
            }

            return result;
        }

        private static string SanitizeLayoutName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "Layout";

            var invalidChars = Path.GetInvalidFileNameChars().Concat(new[] { ':', ';' }).Distinct().ToArray();
            var sanitized = new string(value.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "Layout" : sanitized;
        }

        private bool ConfirmPreflight(SheetPreflightResult result)
        {
            if (result == null || !result.HasIssues)
            {
                StatusMessage = "Preflight passed.";
                return true;
            }

            StatusMessage = result.Summary;
            var message = BuildPreflightMessage(result);

            if (result.HasErrors)
            {
                MessageBox.Show(message, "Preflight Errors", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (result.HasWarnings)
            {
                var choice = MessageBox.Show(
                    message + "\n\nContinue export?",
                    "Preflight Warnings",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                return choice == MessageBoxResult.Yes;
            }

            return true;
        }

        private string BuildPreflightMessage(SheetPreflightResult result)
        {
            var lines = new List<string> { result.Summary, "" };
            var visibleIssues = result.Issues
                .Where(i => i.Severity != PreflightSeverity.Info)
                .Take(16)
                .ToList();

            foreach (var issue in visibleIssues)
                lines.Add(string.Format("[{0}] {1}: {2}", issue.Severity, issue.DisplayName, issue.Message));

            var hiddenCount = result.Issues.Count(i => i.Severity != PreflightSeverity.Info) - visibleIssues.Count;
            if (hiddenCount > 0)
                lines.Add(string.Format("...and {0} more issue(s).", hiddenCount));

            return string.Join(Environment.NewLine, lines);
        }

        private void CancelExport()
        {
            _cancellationTokenSource?.Cancel();
            StatusMessage = "Cancelling...";
        }

        private ExportSettings BuildExportSettings()
        {
            var sortMode = SortMode.SheetNumber;
            if (SelectedSortMode == "Name") sortMode = SortMode.Name;
            else if (SelectedSortMode == "Custom") sortMode = SortMode.Custom;
            else if (SelectedSortMode == "Revit Sheet Schedule") sortMode = SortMode.RevitSheetSchedule;

            var verticalAlign = Models.VerticalAlignment.Top;
            if (SelectedVerticalAlignment == "Center") verticalAlign = Models.VerticalAlignment.Center;
            else if (SelectedVerticalAlignment == "Bottom") verticalAlign = Models.VerticalAlignment.Bottom;

            return new ExportSettings
            {
                OutputFolder = GetResolvedOutputFolder(),
                FileNameTemplate = FileNameTemplate,
                DwgExportSetupName = SelectedSetup,
                ExportMode = ExportMode,
                DwgVersion = SelectedDwgVersion,
                SmartViewScale = SmartViewScale,
                OpenAfterExport = OpenAfterExport,
                OrderRuleSource = SelectedSortMode,
                SelectedSheetScheduleId = SelectedSheetSchedule?.ElementIdValue ?? "",
                VerticalAlign = verticalAlign,
                SortMode = sortMode,
                PreserveCoincidentLines = PreserveCoincidentLines,
                MergeLayers = MergeLayers,
                LayoutNameTemplate = LayoutNameTemplate
            };
        }

        private async Task OpenWithAutoCADAsync(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Trace.WriteLine("[Export] File not found for open: " + filePath);
                MessageBox.Show("File not found:\n" + filePath, "Open Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await Task.Delay(1000);

            try
            {
                var acadPath = AutoCadLocatorService.FindAutoCAD(SelectedAutoCADVersion);
                if (!string.IsNullOrEmpty(acadPath) && File.Exists(acadPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = acadPath,
                        Arguments = "\"" + filePath + "\"",
                        UseShellExecute = true
                    });
                    Trace.WriteLine("[Export] Opened with AutoCAD: " + acadPath);
                    return;
                }

                var accorePath = AutoCadLocatorService.FindAcCoreConsole(SelectedAutoCADVersion);
                if (!string.IsNullOrEmpty(accorePath))
                {
                    var acadDir = Path.GetDirectoryName(accorePath);
                    var acadExe = Path.Combine(acadDir, "acad.exe");
                    if (File.Exists(acadExe))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = acadExe,
                            Arguments = "\"" + filePath + "\"",
                            UseShellExecute = true
                        });
                        Trace.WriteLine("[Export] Opened with AutoCAD: " + acadExe);
                        return;
                    }
                }

                try
                {
                    Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                    Trace.WriteLine("[Export] Opened with default handler");
                }
                catch (Exception openEx)
                {
                    Trace.WriteLine("[Export] No handler for .dwg: " + openEx.Message);
                    MessageBox.Show(
                        "Cannot open DWG file. No AutoCAD installation found.\n\nFile: " + filePath,
                        "Open Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Export] Failed to open file: " + ex.Message);
                MessageBox.Show(
                    "Failed to open DWG file:\n" + ex.Message + "\n\nFile: " + filePath,
                    "Open Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
