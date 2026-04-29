using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.Services;
using Licorp_CombineCAD.Views;

namespace Licorp_CombineCAD.ViewModels
{
    public class ExportDialogViewModel : INotifyPropertyChanged
    {
        private readonly UIDocument _uiDocument;
        private readonly Document _document;
        private readonly SheetCollectorService _sheetCollector;
        private readonly DwgExportService _exportService;
        private readonly SheetPreflightService _preflightService;
        private readonly ProfileService _profileService;
        private readonly ExportProfile _profile;

        private string _filterText = "";
        private string _selectedSetup;
        private string _outputFolder;
        private ExportMode _exportMode = ExportMode.MultiLayout;
        private bool _autoBindXRef = true;
        private bool _smartViewScale = false;
        private bool _openAfterExport = false;
        private bool _progressAlwaysOnTop = true;
        private bool _isExporting = false;
        private int _selectedCount = 0;
        private CancellationTokenSource _cancellationTokenSource;
        private string _statusMessage = "";

        private RevitThreadService _revitThreadService;
        private ProgressDialog _progressDialog;
        private ProgressViewModel _progressVm;

        public ExportDialogViewModel(UIDocument uiDocument)
        {
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _document = uiDocument.Document;

            _sheetCollector = new SheetCollectorService(_document);
            _exportService = new DwgExportService(_document);
            _preflightService = new SheetPreflightService(_document);
            _profileService = new ProfileService();
            _profile = _profileService.LoadLastProfile();

            _revitThreadService = new RevitThreadService();

            LoadSheets();
            LoadProfileSettings();

            var setups = _exportService.GetAvailableExportSetups();
            AvailableSetups = new ObservableCollection<string>(setups);

            SelectAllCommand = new RelayCommand(() => SelectAll(), () => !IsExporting);
            DeselectAllCommand = new RelayCommand(() => DeselectAll(), () => !IsExporting);
            BrowseFolderCommand = new RelayCommand(() => BrowseFolder());
            ExportCommand = new RelayCommand(async () => await ExecuteExportAsync(), CanExport);
            CancelExportCommand = new RelayCommand(() => CancelExport(), () => IsExporting);
            SaveSettingsCommand = new RelayCommand(() => SaveSettings());
            LoadSettingsCommand = new RelayCommand(() => LoadProfileSettings());

            UpdateSelectedCount();
            CheckAutoCADAvailability();
        }

        public ObservableCollection<SheetItemViewModel> AllSheets { get; } = new ObservableCollection<SheetItemViewModel>();

        public ObservableCollection<SheetItemViewModel> FilteredSheets
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_filterText))
                    return AllSheets;

                var filterLower = _filterText.ToLowerInvariant();
                var filtered = AllSheets.Where(s =>
                    s.SheetNumber.ToLowerInvariant().Contains(filterLower) ||
                    s.SheetName.ToLowerInvariant().Contains(filterLower) ||
                    s.Revision.ToLowerInvariant().Contains(filterLower));

                return new ObservableCollection<SheetItemViewModel>(filtered);
            }
        }

        public string FilterText
        {
            get => _filterText;
            set { _filterText = value; OnPropertyChanged(); OnPropertyChanged(nameof(FilteredSheets)); }
        }

        public string OutputFolder
        {
            get => _outputFolder;
            set { _outputFolder = value; OnPropertyChanged(); }
        }

        public ExportMode ExportMode
        {
            get => _exportMode;
            set
            {
                _exportMode = value;
                EnforceSelfContainedSheetDwgExport();
                OnPropertyChanged();
                ValidateExportMode();
            }
        }

        public bool AutoBindXRef
        {
            get => _autoBindXRef;
            set
            {
                _autoBindXRef = value;
                EnforceSelfContainedSheetDwgExport();
                OnPropertyChanged();
            }
        }

        public bool SmartViewScale
        {
            get => _smartViewScale;
            set { _smartViewScale = value; OnPropertyChanged(); }
        }

        public bool OpenAfterExport
        {
            get => _openAfterExport;
            set { _openAfterExport = value; OnPropertyChanged(); }
        }

        public bool ProgressAlwaysOnTop
        {
            get => _progressAlwaysOnTop;
            set { _progressAlwaysOnTop = value; OnPropertyChanged(); }
        }

        public bool IsExporting
        {
            get => _isExporting;
            set { _isExporting = value; OnPropertyChanged(); (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged(); }
        }

        public int SelectedCount
        {
            get => _selectedCount;
            set { _selectedCount = value; OnPropertyChanged(); }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public bool IsAutoCADAvailable { get; private set; }
        public string AutoCADVersion { get; private set; }
        public string AutoCADPath { get; private set; }

        public ObservableCollection<string> AvailableSetups { get; }
        public string SelectedSetup
        {
            get => _selectedSetup;
            set { _selectedSetup = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> AvailableAutoCADVersions { get; } = new ObservableCollection<string>();
        public string SelectedAutoCADVersion { get; set; }

        public ObservableCollection<string> AvailableDwgVersions { get; } = new ObservableCollection<string> { "2018", "2013", "2010", "2007" };

        private string _selectedDwgVersion = "2018";
        public string SelectedDwgVersion
        {
            get => _selectedDwgVersion;
            set { _selectedDwgVersion = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> AvailableSortModes { get; } = new ObservableCollection<string> { "Sheet Number", "Name", "Custom" };
        private string _selectedSortMode = "Sheet Number";
        public string SelectedSortMode
        {
            get => _selectedSortMode;
            set { _selectedSortMode = value; OnPropertyChanged(); ApplySort(); OnPropertyChanged(nameof(FilteredSheets)); }
        }

        public ObservableCollection<string> AvailableVerticalAlignments { get; } = new ObservableCollection<string> { "Top", "Center", "Bottom" };
        private string _selectedVerticalAlignment = "Top";
        public string SelectedVerticalAlignment
        {
            get => _selectedVerticalAlignment;
            set { _selectedVerticalAlignment = value; OnPropertyChanged(); }
        }

        private string _fileNameTemplate = "{SheetNumber} - {SheetName}";
        public string FileNameTemplate
        {
            get => _fileNameTemplate;
            set { _fileNameTemplate = value; OnPropertyChanged(); }
}

    private bool _preserveCoincidentLines;
        public bool PreserveCoincidentLines
        {
            get => _preserveCoincidentLines;
            set { _preserveCoincidentLines = value; OnPropertyChanged(); }
        }

        private bool _createSubfolders;
        public bool CreateSubfolders
        {
            get => _createSubfolders;
            set { _createSubfolders = value; OnPropertyChanged(); }
        }

        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }
        public ICommand BrowseFolderCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand CancelExportCommand { get; }
        public ICommand SaveSettingsCommand { get; }
        public ICommand LoadSettingsCommand { get; }

        private void LoadSheets()
        {
            var sheets = _sheetCollector.GetAllSheets();
            AllSheets.Clear();
            foreach (var sheet in sheets)
            {
                var vm = new SheetItemViewModel(sheet);
                vm.PropertyChanged += OnSheetSelectionChanged;
                AllSheets.Add(vm);
            }
            ApplySort();
        }

        public void ReorderSheet(SheetItemViewModel source, SheetItemViewModel target)
        {
            var sourceIndex = AllSheets.IndexOf(source);
            var targetIndex = AllSheets.IndexOf(target);

            if (sourceIndex < 0 || targetIndex < 0 || sourceIndex == targetIndex)
                return;

            AllSheets.Move(sourceIndex, targetIndex);
            SelectedSortMode = "Custom";
            OnPropertyChanged(nameof(FilteredSheets));
        }

        private void ApplySort()
        {
            var mode = SelectedSortMode;
            if (mode == "Custom") return;

            var items = AllSheets.ToList();
            AllSheets.Clear();

            if (mode == "Name")
            {
                items = items.OrderBy(s => s.SheetName, new NaturalStringComparer()).ToList();
            }
            else
            {
                items = items.OrderBy(s => s.SheetNumber, new NaturalStringComparer()).ToList();
            }

            foreach (var item in items)
                AllSheets.Add(item);
        }

        private void OnSheetSelectionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SheetItemViewModel.IsSelected))
            {
                UpdateSelectedCount();
                (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        private void UpdateSelectedCount()
        {
            SelectedCount = AllSheets.Count(s => s.IsSelected);
        }

        private void SelectAll()
        {
            AllSheets.ToList().ForEach(s => s.IsSelected = true);
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void DeselectAll()
        {
            AllSheets.ToList().ForEach(s => s.IsSelected = false);
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

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
            if (IsExporting || SelectedCount == 0 || string.IsNullOrEmpty(OutputFolder))
                return false;

            // Export is always available; merge requires AutoCAD
            return true;
        }

        private void ValidateExportMode()
        {
            EnforceSelfContainedSheetDwgExport();

            if (!IsAutoCADAvailable)
            {
                StatusMessage = "AutoCAD not detected. Only individual export will be performed.";
            }
            else
            {
                StatusMessage = "";
            }
        }

        private void CheckAutoCADAvailability()
        {
            var info = AutoCadLocatorService.GetAutoCADInfo();
            IsAutoCADAvailable = info.Available;
            AutoCADVersion = info.Version;
            AutoCADPath = info.Path;

            AvailableAutoCADVersions.Clear();
            foreach (var ver in AutoCadLocatorService.GetInstalledVersions())
            {
                AvailableAutoCADVersions.Add(ver);
            }

            if (!string.IsNullOrEmpty(AutoCADVersion) && !AvailableAutoCADVersions.Contains(AutoCADVersion))
            {
                AvailableAutoCADVersions.Insert(0, AutoCADVersion);
            }

            SelectedAutoCADVersion = AutoCADVersion ?? (AvailableAutoCADVersions.FirstOrDefault());

            OnPropertyChanged(nameof(IsAutoCADAvailable));
            OnPropertyChanged(nameof(AutoCADVersion));
            OnPropertyChanged(nameof(AutoCADPath));
            OnPropertyChanged(nameof(AvailableAutoCADVersions));
            OnPropertyChanged(nameof(SelectedAutoCADVersion));
            ValidateExportMode();
        }

        private async Task ExecuteExportAsync()
        {
            if (IsExporting) return;

            if (!string.IsNullOrEmpty(OutputFolder) && !Directory.Exists(OutputFolder))
            {
                var result = MessageBox.Show(
                    "Output folder does not exist. Create it?",
                    "Folder Not Found", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No) return;
                try { Directory.CreateDirectory(OutputFolder); }
                catch (Exception dirEx)
                {
                    MessageBox.Show($"Cannot create folder: {dirEx.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            var selectedSheets = AllSheets.Where(s => s.IsSelected).Select(s => s.Model).ToList();
            var settings = BuildExportSettings();

            SheetPreflightResult preflightResult;
            try
            {
                preflightResult = await _revitThreadService.RunOnRevitThreadAsync(app =>
                    _preflightService.Analyze(selectedSheets, settings));
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[Preflight] Failed: {ex}");
                MessageBox.Show(
                    $"Preflight check failed:\n{ex.Message}",
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
                        warningMsg += $"\n\nFailed: {string.Join(", ", exportResult.FailedSheets)}";
                    if (exportResult.SkippedSheets.Count > 0)
                        warningMsg += $"\n\nSkipped: {string.Join(", ", exportResult.SkippedSheets)}";
                    await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        MessageBox.Show(warningMsg, "Export Warnings", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }));
                }

                string fileToOpen = null;

                if (exportedFiles.Count > 0)
                {
                    // Only merge if AutoCAD is available
                    if (IsAutoCADAvailable)
                    {
                        _progressVm.UpdatePhase("Merging");

                        var accorePath = AutoCadLocatorService.FindAcCoreConsole(SelectedAutoCADVersion);
                        var mergeService = new DwgMergeService(accorePath);
                        mergeService.SetVerticalAlignment(settings.VerticalAlign.ToString());
                        mergeService.SetDwgVersion(settings.DwgVersion ?? "Current");
                        var exportedSheets = exportResult?.ExportedSheets ?? new List<SheetInfo>();
                        var layoutNames = exportedSheets.Select(s => $"{s.SheetNumber} - {s.SheetName}").ToList();
                        if (layoutNames.Count != exportedFiles.Count)
                            layoutNames = exportedFiles.Select(Path.GetFileNameWithoutExtension).ToList();
                        var outputPath = GetUniqueOutputPath();

                        bool mergeSuccess = false;
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
                            StatusMessage = $"Merged {exportedFiles.Count} files to {Path.GetFileName(outputPath)}";
                            if (OpenAfterExport)
                                fileToOpen = outputPath;
                        }
                        else
                        {
                            StatusMessage = "Merge failed.";
                            if (OpenAfterExport && exportedFiles.Count > 0)
                                fileToOpen = exportedFiles.First();
                        }
                    }
                    else
                    {
                        StatusMessage = $"Exported {exportedFiles.Count} individual DWG files. AutoCAD not available for merge.";
                        if (OpenAfterExport && exportedFiles.Count > 0)
                            fileToOpen = exportedFiles.First();
                    }
                }

                _progressVm.StopTimer();
                _progressVm.Completed = true;
                SaveSettings();

                await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try { _progressDialog.Close(); } catch (Exception closeEx) { Trace.WriteLine($"[ExportDialog] Close dialog error: {closeEx.Message}"); }
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
                StatusMessage = $"Error: {ex.Message}";
                Trace.WriteLine($"[Export] Error: {ex}");
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
                        try { _progressDialog.Close(); } catch (Exception closeEx) { Trace.WriteLine($"[ExportDialog] Close dialog error: {closeEx.Message}"); }
                    }));
                }
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        private string GetUniqueOutputPath()
        {
            var baseName = $"Combined_{DateTime.Now:yyyyMMdd_HHmmss}";
            var path = Path.Combine(OutputFolder, baseName + ".dwg");

            int counter = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(OutputFolder, $"{baseName}_{counter}.dwg");
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
                MessageBox.Show(
                    message,
                    "Preflight Errors",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
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

            var verticalAlign = Models.VerticalAlignment.Top;
            if (SelectedVerticalAlignment == "Center") verticalAlign = Models.VerticalAlignment.Center;
            else if (SelectedVerticalAlignment == "Bottom") verticalAlign = Models.VerticalAlignment.Bottom;

            return new ExportSettings
            {
                OutputFolder = OutputFolder,
                FileNameTemplate = FileNameTemplate,
                DwgExportSetupName = SelectedSetup,
                ExportMode = ExportMode,
                DwgVersion = SelectedDwgVersion,
                AutoBindXRef = true,
                SmartViewScale = SmartViewScale,
                OpenAfterExport = OpenAfterExport,
                ProgressAlwaysOnTop = ProgressAlwaysOnTop,
                VerticalAlign = verticalAlign,
                SortMode = sortMode,
                PreserveCoincidentLines = PreserveCoincidentLines,
                CreateSubfolders = CreateSubfolders
            };
        }

        private void EnforceSelfContainedSheetDwgExport()
        {
            // The AutoCAD merge engine expects each exported sheet DWG to be self-contained.
            if (!_autoBindXRef)
            {
                _autoBindXRef = true;
                OnPropertyChanged(nameof(AutoBindXRef));
            }
        }

private void SaveSettings()
    {
        _profile.OutputFolder = OutputFolder;
        _profile.SelectedSetup = SelectedSetup;
        _profile.ExportMode = ExportMode.ToString();
        _profile.DwgVersion = SelectedDwgVersion;
        _profile.FileNameTemplate = FileNameTemplate;
        _profile.AutoBindXRef = AutoBindXRef;
        _profile.SmartViewScale = SmartViewScale;
        _profile.OpenAfterExport = OpenAfterExport;
        _profile.PreserveCoincidentLines = PreserveCoincidentLines;
        _profile.CreateSubfolders = CreateSubfolders;
        _profile.SortMode = SelectedSortMode;
        _profile.VerticalAlign = SelectedVerticalAlignment;
        _profile.LastUsed = DateTime.Now;
        _profileService.SaveLastProfile(_profile);
    }

        private void LoadProfileSettings()
        {
            OutputFolder = _profile.OutputFolder;
            SelectedSetup = _profile.SelectedSetup;

            if (Enum.TryParse<ExportMode>(_profile.ExportMode, out var mode))
                ExportMode = mode;

            if (!string.IsNullOrEmpty(_profile.DwgVersion))
                SelectedDwgVersion = _profile.DwgVersion;

            if (!string.IsNullOrEmpty(_profile.FileNameTemplate))
                FileNameTemplate = _profile.FileNameTemplate;

AutoBindXRef = _profile.AutoBindXRef;
        SmartViewScale = _profile.SmartViewScale;
        OpenAfterExport = _profile.OpenAfterExport;
        PreserveCoincidentLines = _profile.PreserveCoincidentLines;
        CreateSubfolders = _profile.CreateSubfolders;
        if (!string.IsNullOrEmpty(_profile.SortMode) && AvailableSortModes.Contains(_profile.SortMode))
            SelectedSortMode = _profile.SortMode;
        if (!string.IsNullOrEmpty(_profile.VerticalAlign) && AvailableVerticalAlignments.Contains(_profile.VerticalAlign))
            SelectedVerticalAlignment = _profile.VerticalAlign;
    }

    private async Task OpenWithAutoCADAsync(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
        {
            Trace.WriteLine($"[Export] File not found for open: {filePath}");
            MessageBox.Show($"File not found:\n{filePath}", "Open Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = true
                });
                Trace.WriteLine($"[Export] Opened with AutoCAD: {acadPath}");
                return;
            }

            var accorePath = AutoCadLocatorService.FindAcCoreConsole(SelectedAutoCADVersion);
            if (!string.IsNullOrEmpty(accorePath))
            {
                var acadDir = System.IO.Path.GetDirectoryName(accorePath);
                var acadExe = System.IO.Path.Combine(acadDir, "acad.exe");
                if (System.IO.File.Exists(acadExe))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = acadExe,
                        Arguments = $"\"{filePath}\"",
                        UseShellExecute = true
                    });
                    Trace.WriteLine($"[Export] Opened with AutoCAD: {acadExe}");
                    return;
                }
            }

            try
            {
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                Trace.WriteLine($"[Export] Opened with default handler");
            }
            catch (Exception openEx)
            {
                Trace.WriteLine($"[Export] No handler for .dwg: {openEx.Message}");
                MessageBox.Show(
                    $"Cannot open DWG file. No AutoCAD installation found.\n\nFile: {filePath}",
                    "Open Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[Export] Failed to open file: {ex.Message}");
            MessageBox.Show(
                $"Failed to open DWG file:\n{ex.Message}\n\nFile: {filePath}",
                "Open Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke() ?? true;
        public void Execute(object parameter) => _execute();

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }

    public class NaturalStringComparer : IComparer<string>
    {
        private static readonly Regex _numberRegex = new Regex(@"(\d+)", RegexOptions.Compiled);

        public int Compare(string x, string y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            var xParts = _numberRegex.Split(x);
            var yParts = _numberRegex.Split(y);

            for (int i = 0; i < System.Math.Min(xParts.Length, yParts.Length); i++)
            {
                if (int.TryParse(xParts[i], out var xNum) && int.TryParse(yParts[i], out var yNum))
                {
                    var cmp = xNum.CompareTo(yNum);
                    if (cmp != 0) return cmp;
                }
                else
                {
                    var cmp = string.Compare(xParts[i], yParts[i], StringComparison.OrdinalIgnoreCase);
                    if (cmp != 0) return cmp;
                }
            }

            return xParts.Length.CompareTo(yParts.Length);
        }
    }
}
