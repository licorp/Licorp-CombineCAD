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
                MessageBox.Show(
                    "Preflight check failed:\n" + ex.Message,
                    "Preflight Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (!ConfirmPreflight(preflightResult))
                return;

            IsExporting = true;
            _cancellationTokenSource = new CancellationTokenSource();

            _progressDialog = new ProgressDialog { Topmost = ProgressAlwaysOnTop };
            _progressVm = new ProgressViewModel(() => _cancellationTokenSource?.Cancel());
            _progressDialog.DataContext = _progressVm;
            _progressVm.StartTimer();
            _progressDialog.Show();

            try
            {
                var options = _exportService.BuildExportOptions(settings);
                var cts = _cancellationTokenSource;
                var progressVm = _progressVm;

                ExportResult exportResult = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    var progress = new DirectProgress<ExportProgressInfo>(info =>
                    {
                        Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            progressVm.Update(info.Phase, info.CurrentItem, info.Current, info.Total);
                        }));
                    });

                    return _exportService.ExportSheetsIndividually(
                        selectedSheets, settings, options,
                        progress,
                        cts.Token);
                });

                var exportedFiles = exportResult?.ExportedFiles ?? new List<string>();

                if (cts.Token.IsCancellationRequested)
                {
                    StatusMessage = "Export cancelled.";
                    _progressVm.StopTimer();
                    _progressVm.Completed = true;
                    return;
                }

                if (exportResult != null && exportResult.HasWarnings)
                {
                    var warningMsg = exportResult.Summary;
                    if (exportResult.FailedSheets.Count > 0)
                        warningMsg += "\n\nFailed: " + string.Join(", ", exportResult.FailedSheets);
                    if (exportResult.SkippedSheets.Count > 0)
                        warningMsg += "\n\nSkipped: " + string.Join(", ", exportResult.SkippedSheets);

                    await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        MessageBox.Show(warningMsg, "Export Warnings", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }));
                }

                string fileToOpen = null;

                if (exportedFiles.Count > 0)
                {
                    if (IsAutoCADAvailable)
                    {
                        _progressVm.UpdatePhase("Merging");

                        var accorePath = AutoCadLocatorService.FindAcCoreConsole(SelectedAutoCADVersion);
                        var acadPath = AutoCadLocatorService.FindAutoCAD(SelectedAutoCADVersion);
                        var mergeService = new DwgMergeService(accorePath, acadPath);
                        mergeService.SetVerticalAlignment(settings.VerticalAlign.ToString());
                        mergeService.SetDwgVersion(settings.DwgVersion ?? "Current");
                        mergeService.SetExpectedSheetCount(exportedFiles.Count);

                        var exportedSheets = exportResult?.ExportedSheets ?? new List<SheetInfo>();
                        var layoutNames = exportedSheets.Select(s => s.SheetNumber + " - " + s.SheetName).ToList();
                        if (layoutNames.Count != exportedFiles.Count)
                            layoutNames = exportedFiles.Select(Path.GetFileNameWithoutExtension).ToList();

                        var outputPath = GetUniqueOutputPath();
                        var mergeSuccess = false;

                        switch (ExportMode)
                        {
                            case ExportMode.MultiLayout:
                                mergeSuccess = await mergeService.MergeToMultiLayoutAsync(exportedFiles, outputPath, layoutNames, null, cts.Token);
                                break;
                            case ExportMode.SingleLayout:
                                mergeSuccess = await mergeService.MergeToSingleLayoutAsync(exportedFiles, outputPath, "Combined", null, cts.Token);
                                break;
                            case ExportMode.ModelSpace:
                                mergeSuccess = await mergeService.MergeToModelSpaceAsync(exportedFiles, outputPath, layoutNames, null, cts.Token);
                                break;
                        }

                        if (mergeSuccess)
                        {
                            StatusMessage = string.Format("Merged {0} files to {1}", exportedFiles.Count, Path.GetFileName(outputPath));
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

                            await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
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
                }

                _progressVm.StopTimer();
                _progressVm.Completed = true;
                SaveSettings();

                await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try { _progressDialog.Close(); }
                    catch (Exception closeEx) { Trace.WriteLine("[ExportDialog] Close dialog error: " + closeEx.Message); }
                }));
                await Task.Delay(500);

                if (fileToOpen != null)
                    await OpenWithAutoCADAsync(fileToOpen);
            }
            catch (OperationCanceledException)
            {
                StatusMessage = "Export cancelled.";
                Trace.WriteLine("[Export] Cancelled by user");
            }
            catch (Exception ex)
            {
                StatusMessage = "Error: " + ex.Message;
                Trace.WriteLine("[Export] Error: " + ex);
            }
            finally
            {
                IsExporting = false;
                if (_progressVm != null && !_progressVm.Completed)
                {
                    _progressVm.StopTimer();
                    _progressVm.Completed = true;
                    await Task.Delay(500);
                    await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try { _progressDialog.Close(); }
                        catch (Exception closeEx) { Trace.WriteLine("[ExportDialog] Close dialog error: " + closeEx.Message); }
                    }));
                }

                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                RefreshDerivedState();
            }
        }

        private string GetUniqueOutputPath()
        {
            var baseName = BuildProjectFileBaseName();
            var path = Path.Combine(GetResolvedOutputFolder(), baseName + ".dwg");

            var counter = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(GetResolvedOutputFolder(), baseName + "_" + counter + ".dwg");
                counter++;
            }

            return path;
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
                ProgressAlwaysOnTop = ProgressAlwaysOnTop,
                OrderRuleSource = SelectedSortMode,
                SelectedSheetScheduleId = SelectedSheetSchedule?.ElementIdValue ?? "",
                VerticalAlign = verticalAlign,
                SortMode = sortMode,
                PreserveCoincidentLines = PreserveCoincidentLines
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
