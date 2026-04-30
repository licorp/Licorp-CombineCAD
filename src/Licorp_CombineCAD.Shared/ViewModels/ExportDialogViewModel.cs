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
        private readonly SheetScheduleService _sheetScheduleService;
        private readonly ViewSheetSetService _viewSheetSetService;
        private readonly ProfileService _profileService;
        private readonly RevitThreadService _revitThreadService;

        private ExportProfile _profile;
        private ExportProfile _selectedProfile;
        private bool _isInitialized;
        private bool _isClosing;
        private bool _isLoadingSheets = true;
        private bool _isExporting;
        private bool _isUpdatingSetSelection;
        private string _loadingMessage = "Opening form...";
        private string _filterText = "";
        private string _selectedSetup;
        private string _selectedAutoCADVersion;
        private string _outputFolder;
        private string _selectedDwgVersion = "2018";
        private string _selectedSortMode = "Sheet Number";
        private string _selectedVerticalAlignment = "Top";
        private string _fileNameTemplate = "{SheetNumber} - {SheetName}";
        private string _selectedMergeEngine = "AcCoreConsole";
        private string _selectedRasterImageMode = "Keep Reference";
        private string _statusMessage = "";
        private string _projectName = "Project";
        private string _projectNumber = "PRJ-001";
        private ExportMode _exportMode = ExportMode.MultiLayout;
        private bool _autoBindXRef = true;
        private bool _smartViewScale;
        private bool _openAfterExport;
        private bool _progressAlwaysOnTop = true;
        private bool _createSheetSet = true;
        private bool _verifyCombinedDwg = true;
        private bool _preserveCoincidentLines;
        private bool _createSubfolders;
        private int _selectedCount;
        private SheetScheduleInfo _selectedSheetSchedule;
        private CancellationTokenSource _cancellationTokenSource;
        private ProgressDialog _progressDialog;
        private ProgressViewModel _progressVm;

        public ExportDialogViewModel(UIDocument uiDocument)
        {
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _document = uiDocument.Document;

            _sheetCollector = new SheetCollectorService(_document);
            _exportService = new DwgExportService(_document);
            _preflightService = new SheetPreflightService(_document);
            _sheetScheduleService = new SheetScheduleService(_document);
            _viewSheetSetService = new ViewSheetSetService(_document);
            _profileService = new ProfileService();
            _revitThreadService = new RevitThreadService();

            Profiles = new ObservableCollection<ExportProfile>(_profileService.LoadProfiles());
            _profile = _profileService.LoadLastProfile();
            SelectedProfile = Profiles.FirstOrDefault(p => p.Id == _profile.Id) ?? Profiles.FirstOrDefault();

            ApplyProfileValues(_profile);
            CheckAutoCADAvailability();

            SelectAllCommand = new RelayCommand(() => SelectVisibleSheets(true), () => !IsExporting && !IsLoadingSheets);
            DeselectAllCommand = new RelayCommand(() => SelectVisibleSheets(false), () => !IsExporting && !IsLoadingSheets);
            BrowseFolderCommand = new RelayCommand(() => BrowseFolder(), () => !IsExporting);
            ExportCommand = new RelayCommand(async () => await ExecuteExportAsync(), CanExport);
            CancelExportCommand = new RelayCommand(() => CancelExport(), () => IsExporting);
            ApplyProfileCommand = new RelayCommand(() => ApplySelectedProfile(), () => SelectedProfile != null && !IsExporting);
            NewProfileCommand = new RelayCommand(() => CreateProfileFromCurrentSettings(), () => !IsExporting);
            SaveProfileCommand = new RelayCommand(() => SaveSettings(), () => SelectedProfile != null && !IsExporting);
            DeleteProfileCommand = new RelayCommand(() => DeleteSelectedProfile(), () => SelectedProfile != null && Profiles.Count > 1 && !IsExporting);
            EditFileNameTemplateCommand = new RelayCommand(() => EditFileNameTemplate(), () => !IsExporting);
            ReorderSelectedSheetsCommand = new RelayCommand(() => ReorderSelectedSheets(), () => SelectedCount > 1 && !IsExporting && !IsLoadingSheets);
            SaveSheetSetCommand = new RelayCommand(async () => await SaveSelectedSheetsAsSetAsync(), () => SelectedCount > 0 && !IsExporting && !IsLoadingSheets);
            LoadSettingsCommand = new RelayCommand(() => ApplySelectedProfile(), () => SelectedProfile != null && !IsExporting);
            SaveSettingsCommand = SaveProfileCommand;

            RefreshDerivedState();
        }

        public ObservableCollection<ExportProfile> Profiles { get; }
        public ObservableCollection<SheetItemViewModel> AllSheets { get; } = new ObservableCollection<SheetItemViewModel>();
        public ObservableCollection<ViewSheetSetInfo> AvailableViewSheetSets { get; } = new ObservableCollection<ViewSheetSetInfo>();
        public ObservableCollection<string> AvailableSetups { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> AvailableAutoCADVersions { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> AvailableDwgVersions { get; } = new ObservableCollection<string> { "2018", "2013", "2010", "2007" };
        public ObservableCollection<string> AvailableSortModes { get; } = new ObservableCollection<string> { "Sheet Number", "Name", "Custom", "Revit Sheet Schedule" };
        public ObservableCollection<SheetScheduleInfo> AvailableSheetSchedules { get; } = new ObservableCollection<SheetScheduleInfo>();
        public ObservableCollection<string> AvailableVerticalAlignments { get; } = new ObservableCollection<string> { "Top", "Center", "Bottom" };
        public ObservableCollection<string> AvailableMergeEngines { get; } = new ObservableCollection<string> { "AcCoreConsole", "Full AutoCAD" };
        public ObservableCollection<string> AvailableRasterImageModes { get; } = new ObservableCollection<string> { "Keep Reference", "Embed as OLE" };

        public ObservableCollection<SheetItemViewModel> FilteredSheets => new ObservableCollection<SheetItemViewModel>(GetVisibleSheets());

        public ExportProfile SelectedProfile
        {
            get => _selectedProfile;
            set { _selectedProfile = value; OnPropertyChanged(); }
        }

        public bool IsLoadingSheets
        {
            get => _isLoadingSheets;
            set
            {
                _isLoadingSheets = value;
                OnPropertyChanged();
                (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (SelectAllCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (DeselectAllCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (ReorderSelectedSheetsCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (SaveSheetSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        public string LoadingMessage
        {
            get => _loadingMessage;
            set { _loadingMessage = value; OnPropertyChanged(); }
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                _filterText = value ?? "";
                OnPropertyChanged();
                RefreshFilterState();
            }
        }

        public string OutputFolder
        {
            get => _outputFolder;
            set { _outputFolder = value; OnPropertyChanged(); RefreshDerivedState(); }
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
                RefreshDerivedState();
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

        public bool CreateSheetSet
        {
            get => _createSheetSet;
            set { _createSheetSet = value; OnPropertyChanged(); }
        }

        public bool VerifyCombinedDwg
        {
            get => _verifyCombinedDwg;
            set { _verifyCombinedDwg = value; OnPropertyChanged(); }
        }

        public bool IsExporting
        {
            get => _isExporting;
            set
            {
                _isExporting = value;
                OnPropertyChanged();
                (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (CancelExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (SaveSheetSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        public int SelectedCount
        {
            get => _selectedCount;
            set { _selectedCount = value; OnPropertyChanged(); }
        }

        public int VisibleCount => GetVisibleSheets().Count;
        public int TotalCount => AllSheets.Count;
        public string SelectionSummary => string.Format("{0} selected / {1} visible / {2} total", SelectedCount, VisibleCount, TotalCount);

        public bool AreVisibleSheetsSelected
        {
            get
            {
                var visible = GetVisibleSheets();
                return visible.Count > 0 && visible.All(s => s.IsSelected);
            }
            set
            {
                SelectVisibleSheets(value);
                OnPropertyChanged();
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public bool IsAutoCADAvailable { get; private set; }
        public string AutoCADVersion { get; private set; }
        public string AutoCADPath { get; private set; }
        public string AutoCADStatusText => IsAutoCADAvailable
            ? string.Format("Detected {0} - {1}", AutoCADVersion, SelectedMergeEngine)
            : "AutoCAD not detected - individual DWG export only";

        public string SelectedSetup
        {
            get => _selectedSetup;
            set { _selectedSetup = value; OnPropertyChanged(); RefreshDerivedState(); }
        }

        public string SelectedAutoCADVersion
        {
            get => _selectedAutoCADVersion;
            set { _selectedAutoCADVersion = value; OnPropertyChanged(); RefreshDerivedState(); }
        }

        public string SelectedDwgVersion
        {
            get => _selectedDwgVersion;
            set { _selectedDwgVersion = value; OnPropertyChanged(); RefreshDerivedState(); }
        }

        public string SelectedSortMode
        {
            get => _selectedSortMode;
            set
            {
                _selectedSortMode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsSheetScheduleSort));
                ApplySort();
                RefreshFilterState();
            }
        }

        public bool IsSheetScheduleSort => SelectedSortMode == "Revit Sheet Schedule";

        public SheetScheduleInfo SelectedSheetSchedule
        {
            get => _selectedSheetSchedule;
            set
            {
                _selectedSheetSchedule = value;
                OnPropertyChanged();
                if (IsSheetScheduleSort)
                {
                    ApplySort();
                    RefreshFilterState();
                }
            }
        }

        public string SelectedVerticalAlignment
        {
            get => _selectedVerticalAlignment;
            set { _selectedVerticalAlignment = value; OnPropertyChanged(); RefreshDerivedState(); }
        }

        public string FileNameTemplate
        {
            get => _fileNameTemplate;
            set { _fileNameTemplate = value; OnPropertyChanged(); RefreshDerivedState(); }
        }

        public string FileNamePreview => BuildFileNamePreview();

        public bool PreserveCoincidentLines
        {
            get => _preserveCoincidentLines;
            set { _preserveCoincidentLines = value; OnPropertyChanged(); }
        }

        public bool CreateSubfolders
        {
            get => _createSubfolders;
            set { _createSubfolders = value; OnPropertyChanged(); }
        }

        public string SelectedMergeEngine
        {
            get => _selectedMergeEngine;
            set { _selectedMergeEngine = value; OnPropertyChanged(); RefreshDerivedState(); }
        }

        public string SelectedRasterImageMode
        {
            get => _selectedRasterImageMode;
            set { _selectedRasterImageMode = value; OnPropertyChanged(); }
        }

        public string SelectedSetsDisplay
        {
            get
            {
                var selected = AvailableViewSheetSets.Where(s => s.IsSelected).ToList();
                if (selected.Count == 0 || selected.Any(s => s.IsBuiltIn))
                    return "All Sheets";
                if (selected.Count <= 2)
                    return string.Join(", ", selected.Select(s => s.Name));
                return selected.Count + " sets selected";
            }
        }

        public string ExportSummary => string.Format(
            "{0} | {1} sheet(s) | Sort: {2}",
            ExportMode,
            SelectedCount,
            SelectedSortMode);

        public string CombinedFilePreview => string.IsNullOrWhiteSpace(OutputFolder)
            ? "Choose output folder"
            : Path.Combine(OutputFolder, "Combined_yyyyMMdd_HHmmss.dwg");

        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }
        public ICommand BrowseFolderCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand CancelExportCommand { get; }
        public ICommand ApplyProfileCommand { get; }
        public ICommand NewProfileCommand { get; }
        public ICommand SaveProfileCommand { get; }
        public ICommand DeleteProfileCommand { get; }
        public ICommand EditFileNameTemplateCommand { get; }
        public ICommand ReorderSelectedSheetsCommand { get; }
        public ICommand SaveSheetSetCommand { get; }
        public ICommand SaveSettingsCommand { get; }
        public ICommand LoadSettingsCommand { get; }

        public async Task InitializeAsync()
        {
            if (_isInitialized)
                return;

            _isInitialized = true;
            IsLoadingSheets = true;
            LoadingMessage = "Loading sheet number, sheet name, and sheet sets...";

            try
            {
                var data = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    return new InitialLoadData
                    {
                        Sheets = _sheetCollector.GetAllSheets(),
                        Schedules = _sheetScheduleService.GetSheetListSchedules(),
                        ViewSheetSets = _viewSheetSetService.GetSheetSets(),
                        ExportSetups = _exportService.GetAvailableExportSetups(),
                        ProjectName = GetProjectInfoValue("Name", "Project"),
                        ProjectNumber = GetProjectInfoValue("Number", "PRJ-001")
                    };
                });

                if (_isClosing)
                    return;

                _projectName = data.ProjectName;
                _projectNumber = data.ProjectNumber;
                LoadExportSetups(data.ExportSetups);
                LoadSheetSchedules(data.Schedules);
                LoadViewSheetSets(data.ViewSheetSets);
                LoadSheets(data.Sheets);
                ApplySavedScheduleSelection();
                StatusMessage = "Ready.";
            }
            catch (Exception ex)
            {
                StatusMessage = "Failed to load sheets: " + ex.Message;
                Trace.WriteLine("[ExportDialog] Load failed: " + ex);
            }
            finally
            {
                IsLoadingSheets = false;
                LoadingMessage = "";
                RefreshDerivedState();
            }
        }

        public void Dispose()
        {
            _isClosing = true;
            try { _cancellationTokenSource?.Cancel(); } catch { }
            try { _revitThreadService?.Dispose(); } catch { }
        }

        public void ReorderSheet(SheetItemViewModel source, SheetItemViewModel target)
        {
            var sourceIndex = AllSheets.IndexOf(source);
            var targetIndex = AllSheets.IndexOf(target);
            if (sourceIndex < 0 || targetIndex < 0 || sourceIndex == targetIndex)
                return;

            AllSheets.Move(sourceIndex, targetIndex);
            SelectedSortMode = "Custom";
            RefreshFilterState();
        }

        private void LoadExportSetups(IEnumerable<string> setups)
        {
            AvailableSetups.Clear();
            foreach (var setup in setups ?? Enumerable.Empty<string>())
                AvailableSetups.Add(setup);

            if (string.IsNullOrWhiteSpace(SelectedSetup) && AvailableSetups.Count > 0)
                SelectedSetup = AvailableSetups[0];
        }

        private void LoadSheets(IEnumerable<SheetInfo> sheets)
        {
            foreach (var old in AllSheets)
                old.PropertyChanged -= OnSheetSelectionChanged;

            AllSheets.Clear();
            foreach (var sheet in sheets ?? Enumerable.Empty<SheetInfo>())
            {
                var vm = new SheetItemViewModel(sheet);
                vm.PropertyChanged += OnSheetSelectionChanged;
                AllSheets.Add(vm);
            }

            ApplySort();
            UpdateSelectedCount();
            RefreshFilterState();
        }

        private void LoadSheetSchedules(IEnumerable<SheetScheduleInfo> schedules)
        {
            AvailableSheetSchedules.Clear();
            foreach (var schedule in schedules ?? Enumerable.Empty<SheetScheduleInfo>())
                AvailableSheetSchedules.Add(schedule);

            if (SelectedSheetSchedule == null && AvailableSheetSchedules.Count > 0)
                SelectedSheetSchedule = AvailableSheetSchedules[0];
        }

        private void LoadViewSheetSets(IEnumerable<ViewSheetSetInfo> sets)
        {
            foreach (var old in AvailableViewSheetSets)
                old.PropertyChanged -= OnViewSheetSetChanged;

            AvailableViewSheetSets.Clear();
            foreach (var set in sets ?? Enumerable.Empty<ViewSheetSetInfo>())
            {
                set.PropertyChanged += OnViewSheetSetChanged;
                AvailableViewSheetSets.Add(set);
            }

            var allSheets = AvailableViewSheetSets.FirstOrDefault(s => s.IsBuiltIn);
            if (allSheets != null)
                allSheets.IsSelected = true;

            RefreshFilterState();
        }

        private void ApplySavedScheduleSelection()
        {
            if (string.IsNullOrWhiteSpace(_profile?.SelectedSheetScheduleId))
                return;

            var saved = AvailableSheetSchedules.FirstOrDefault(s => s.ElementIdValue == _profile.SelectedSheetScheduleId);
            if (saved != null)
                SelectedSheetSchedule = saved;
        }

        private List<SheetItemViewModel> GetVisibleSheets()
        {
            IEnumerable<SheetItemViewModel> query = AllSheets;
            var selectedSets = AvailableViewSheetSets.Where(s => s.IsSelected).ToList();
            var allSheetsSelected = selectedSets.Count == 0 || selectedSets.Any(s => s.IsBuiltIn);

            if (!allSheetsSelected)
            {
                var allowedIds = new HashSet<string>(
                    selectedSets.SelectMany(s => s.SheetIdValues),
                    StringComparer.OrdinalIgnoreCase);

                query = query.Where(s => allowedIds.Contains(ViewSheetSetService.GetElementIdValue(s.ElementId)));
            }

            if (!string.IsNullOrWhiteSpace(FilterText))
            {
                var filter = FilterText.Trim().ToLowerInvariant();
                query = query.Where(s =>
                    (s.SheetNumber ?? "").ToLowerInvariant().Contains(filter) ||
                    (s.SheetName ?? "").ToLowerInvariant().Contains(filter) ||
                    (s.PaperSize ?? "").ToLowerInvariant().Contains(filter));
            }

            return query.ToList();
        }

        private void OnViewSheetSetChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(ViewSheetSetInfo.IsSelected) || _isUpdatingSetSelection)
                return;

            try
            {
                _isUpdatingSetSelection = true;
                var changed = sender as ViewSheetSetInfo;
                if (changed != null && changed.IsSelected)
                {
                    if (changed.IsBuiltIn)
                    {
                        foreach (var set in AvailableViewSheetSets.Where(s => !s.IsBuiltIn))
                            set.IsSelected = false;
                    }
                    else
                    {
                        foreach (var set in AvailableViewSheetSets.Where(s => s.IsBuiltIn))
                            set.IsSelected = false;
                    }
                }

                if (!AvailableViewSheetSets.Any(s => s.IsSelected))
                {
                    var allSheets = AvailableViewSheetSets.FirstOrDefault(s => s.IsBuiltIn);
                    if (allSheets != null)
                        allSheets.IsSelected = true;
                }
            }
            finally
            {
                _isUpdatingSetSelection = false;
            }

            RefreshFilterState();
        }

        private void ApplySort()
        {
            if (SelectedSortMode == "Custom")
                return;

            if (SelectedSortMode == "Revit Sheet Schedule")
            {
                ApplyRevitScheduleSort();
                return;
            }

            var items = AllSheets.ToList();
            items = SelectedSortMode == "Name"
                ? items.OrderBy(s => s.SheetName, new NaturalStringComparer()).ToList()
                : items.OrderBy(s => s.SheetNumber, new NaturalStringComparer()).ToList();

            ReplaceSheets(items);
        }

        private void ApplyRevitScheduleSort()
        {
            var items = AllSheets.ToList();
            if (items.Count == 0)
                return;

            var orderedNumbers = _sheetScheduleService.GetOrderedSheetNumbers(
                SelectedSheetSchedule?.ElementIdValue,
                items.Select(i => i.Model));

            if (orderedNumbers.Count == 0)
            {
                Trace.WriteLine("[SheetSchedule] No usable schedule order found; falling back to sheet number order");
                items = items.OrderBy(s => s.SheetNumber, new NaturalStringComparer()).ToList();
            }
            else
            {
                var rank = orderedNumbers
                    .Select((number, index) => new { number, index })
                    .GroupBy(x => x.number, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First().index, StringComparer.OrdinalIgnoreCase);

                items = items
                    .OrderBy(s => rank.TryGetValue(s.SheetNumber ?? "", out var order) ? order : int.MaxValue)
                    .ThenBy(s => s.SheetNumber, new NaturalStringComparer())
                    .ToList();
            }

            ReplaceSheets(items);
        }

        private void ReplaceSheets(IEnumerable<SheetItemViewModel> items)
        {
            AllSheets.Clear();
            foreach (var item in items)
                AllSheets.Add(item);
        }

        private void OnSheetSelectionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(SheetItemViewModel.IsSelected))
                return;

            UpdateSelectedCount();
            RefreshDerivedState();
        }

        private void UpdateSelectedCount()
        {
            SelectedCount = AllSheets.Count(s => s.IsSelected);
        }

        private void SelectVisibleSheets(bool isSelected)
        {
            foreach (var sheet in GetVisibleSheets())
                sheet.IsSelected = isSelected;

            UpdateSelectedCount();
            RefreshDerivedState();
        }

        private void ReorderSelectedSheets()
        {
            var selected = AllSheets.Where(s => s.IsSelected).ToList();
            if (selected.Count < 2)
                return;

            var dialog = new ReorderSheetsDialog(selected)
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() != true)
                return;

            var selectedById = selected.ToDictionary(
                s => ViewSheetSetService.GetElementIdValue(s.ElementId),
                s => s,
                StringComparer.OrdinalIgnoreCase);

            var orderedSelected = dialog.OrderedIds
                .Where(id => selectedById.ContainsKey(id))
                .Select(id => selectedById[id])
                .ToList();

            if (orderedSelected.Count != selected.Count)
                return;

            var queue = new Queue<SheetItemViewModel>(orderedSelected);
            var reordered = AllSheets
                .Select(item => item.IsSelected ? queue.Dequeue() : item)
                .ToList();

            ReplaceSheets(reordered);
            SelectedSortMode = "Custom";
            RefreshFilterState();
        }

        private async Task SaveSelectedSheetsAsSetAsync()
        {
            var selected = AllSheets.Where(s => s.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Select at least one sheet before saving a Sheet Set.",
                    "No Sheets Selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var dialog = new SaveSheetSetDialog(selected.Count)
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() != true)
                return;

            var setName = dialog.SetName;
            var selectedIds = selected.Select(s => s.ElementId).ToList();

            try
            {
                var exists = await _revitThreadService.RunOnRevitThreadAsync(app =>
                    _viewSheetSetService.SheetSetExists(setName));

                var replaceExisting = false;
                if (exists)
                {
                    var result = MessageBox.Show(
                        "A Sheet Set named '" + setName + "' already exists.\n\nReplace it?",
                        "Sheet Set Exists",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (result != MessageBoxResult.Yes)
                        return;

                    replaceExisting = true;
                }

                var sets = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    _viewSheetSetService.CreateSheetSet(setName, selectedIds, replaceExisting);
                    return _viewSheetSetService.GetSheetSets();
                });

                LoadViewSheetSets(sets);
                SelectOnlySheetSet(setName);
                StatusMessage = "Saved Sheet Set: " + setName;

                MessageBox.Show(
                    "Sheet Set '" + setName + "' saved successfully.\n\nContains " + selected.Count + " sheet(s).",
                    "Sheet Set Saved",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[ViewSheetSet] Save failed: " + ex);
                MessageBox.Show(
                    "Failed to save Sheet Set:\n" + ex.Message,
                    "Sheet Set Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SelectOnlySheetSet(string setName)
        {
            try
            {
                _isUpdatingSetSelection = true;
                foreach (var set in AvailableViewSheetSets)
                    set.IsSelected = !set.IsBuiltIn && string.Equals(set.Name, setName, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                _isUpdatingSetSelection = false;
            }

            RefreshFilterState();
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
            return !IsExporting &&
                !IsLoadingSheets &&
                SelectedCount > 0 &&
                !string.IsNullOrEmpty(OutputFolder);
        }

        private void ValidateExportMode()
        {
            EnforceSelfContainedSheetDwgExport();
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
                        var mergeService = new DwgMergeService(accorePath, settings.MergeEngine, acadPath);
                        mergeService.SetVerticalAlignment(settings.VerticalAlign.ToString());
                        mergeService.SetDwgVersion(settings.DwgVersion ?? "Current");
                        mergeService.SetReliabilityOptions(
                            settings.VerifyCombinedDwg,
                            settings.CreateSheetSet,
                            settings.RasterImageMode.ToString(),
                            exportedFiles.Count);

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
            var baseName = "Combined_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var path = Path.Combine(OutputFolder, baseName + ".dwg");

            var counter = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(OutputFolder, baseName + "_" + counter + ".dwg");
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

            var mergeEngine = SelectedMergeEngine == "Full AutoCAD"
                ? MergeEngine.FullAutoCAD
                : MergeEngine.AcCoreConsole;

            var rasterImageMode = SelectedRasterImageMode == "Embed as OLE"
                ? RasterImageMode.EmbedAsOle
                : RasterImageMode.KeepReference;

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
                OrderRuleSource = SelectedSortMode,
                SelectedSheetScheduleId = SelectedSheetSchedule?.ElementIdValue ?? "",
                CreateSheetSet = CreateSheetSet,
                RasterImageMode = rasterImageMode,
                MergeEngine = mergeEngine,
                VerifyCombinedDwg = VerifyCombinedDwg,
                VerticalAlign = verticalAlign,
                SortMode = sortMode,
                PreserveCoincidentLines = PreserveCoincidentLines,
                CreateSubfolders = CreateSubfolders
            };
        }

        private void EnforceSelfContainedSheetDwgExport()
        {
            if (!_autoBindXRef)
            {
                _autoBindXRef = true;
                OnPropertyChanged(nameof(AutoBindXRef));
            }
        }

        private void ApplySelectedProfile()
        {
            if (SelectedProfile == null)
                return;

            _profile = SelectedProfile;
            _profileService.SetLastProfile(_profile.Id);
            ApplyProfileValues(_profile);
            StatusMessage = "Profile applied: " + _profile.Name;
        }

        private void CreateProfileFromCurrentSettings()
        {
            var dialog = new ProfileNameDialog("New profile name")
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() != true)
                return;

            if (Profiles.Any(p => string.Equals(p.Name, dialog.ProfileName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Profile already exists: " + dialog.ProfileName, "Profile", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var profile = new ExportProfile
            {
                Id = Guid.NewGuid().ToString("N"),
                Name = dialog.ProfileName,
                Description = "Created " + DateTime.Now.ToString("yyyy-MM-dd HH:mm")
            };

            SaveCurrentSettingsToProfile(profile);
            _profileService.SaveLastProfile(profile);
            Profiles.Add(profile);
            SelectedProfile = profile;
            _profile = profile;
            StatusMessage = "Profile created: " + profile.Name;
        }

        private void DeleteSelectedProfile()
        {
            if (SelectedProfile == null || Profiles.Count <= 1)
                return;

            var confirm = MessageBox.Show(
                "Delete profile '" + SelectedProfile.Name + "'?",
                "Delete Profile",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes)
                return;

            var deleted = SelectedProfile;
            _profileService.DeleteProfile(deleted.Id);
            Profiles.Remove(deleted);
            SelectedProfile = Profiles.FirstOrDefault();
            _profile = SelectedProfile;

            if (_profile != null)
            {
                _profileService.SetLastProfile(_profile.Id);
                ApplyProfileValues(_profile);
            }
        }

        private void SaveSettings()
        {
            if (SelectedProfile == null)
            {
                CreateProfileFromCurrentSettings();
                return;
            }

            _profile = SelectedProfile;
            SaveCurrentSettingsToProfile(_profile);
            _profileService.SaveLastProfile(_profile);
            StatusMessage = "Profile saved: " + _profile.Name;
        }

        private void SaveCurrentSettingsToProfile(ExportProfile profile)
        {
            profile.OutputFolder = OutputFolder;
            profile.SelectedSetup = SelectedSetup;
            profile.ExportMode = ExportMode.ToString();
            profile.DwgVersion = SelectedDwgVersion;
            profile.FileNameTemplate = FileNameTemplate;
            profile.AutoBindXRef = AutoBindXRef;
            profile.SmartViewScale = SmartViewScale;
            profile.OpenAfterExport = OpenAfterExport;
            profile.ProgressAlwaysOnTop = ProgressAlwaysOnTop;
            profile.PreserveCoincidentLines = PreserveCoincidentLines;
            profile.CreateSubfolders = CreateSubfolders;
            profile.SortMode = SelectedSortMode;
            profile.SelectedSheetScheduleId = SelectedSheetSchedule?.ElementIdValue ?? "";
            profile.VerticalAlign = SelectedVerticalAlignment;
            profile.CreateSheetSet = CreateSheetSet;
            profile.RasterImageMode = SelectedRasterImageMode == "Embed as OLE" ? "EmbedAsOle" : "KeepReference";
            profile.MergeEngine = SelectedMergeEngine == "Full AutoCAD" ? "FullAutoCAD" : "AcCoreConsole";
            profile.VerifyCombinedDwg = VerifyCombinedDwg;
            profile.LastUsed = DateTime.Now;
        }

        private void ApplyProfileValues(ExportProfile profile)
        {
            if (profile == null)
                return;

            OutputFolder = profile.OutputFolder;
            SelectedSetup = profile.SelectedSetup;

            if (Enum.TryParse(profile.ExportMode, out ExportMode mode))
                ExportMode = mode;

            if (!string.IsNullOrEmpty(profile.DwgVersion))
                SelectedDwgVersion = profile.DwgVersion;

            if (!string.IsNullOrEmpty(profile.FileNameTemplate))
                FileNameTemplate = profile.FileNameTemplate;

            AutoBindXRef = profile.AutoBindXRef;
            SmartViewScale = profile.SmartViewScale;
            OpenAfterExport = profile.OpenAfterExport;
            ProgressAlwaysOnTop = profile.ProgressAlwaysOnTop;
            PreserveCoincidentLines = profile.PreserveCoincidentLines;
            CreateSubfolders = profile.CreateSubfolders;

            if (!string.IsNullOrEmpty(profile.SortMode) && AvailableSortModes.Contains(profile.SortMode))
                SelectedSortMode = profile.SortMode;

            if (!string.IsNullOrEmpty(profile.VerticalAlign) && AvailableVerticalAlignments.Contains(profile.VerticalAlign))
                SelectedVerticalAlignment = profile.VerticalAlign;

            CreateSheetSet = profile.CreateSheetSet;
            SelectedRasterImageMode = profile.RasterImageMode == "EmbedAsOle" ? "Embed as OLE" : "Keep Reference";
            SelectedMergeEngine = profile.MergeEngine == "FullAutoCAD" ? "Full AutoCAD" : "AcCoreConsole";
            VerifyCombinedDwg = profile.VerifyCombinedDwg;
            ApplySavedScheduleSelection();
            RefreshDerivedState();
        }

        private void EditFileNameTemplate()
        {
            var dialog = new CustomFileNameDialog(FileNameTemplate, _projectNumber, _projectName)
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() == true)
                FileNameTemplate = dialog.FileNameTemplate;
        }

        private string BuildFileNamePreview()
        {
            var sample = AllSheets.FirstOrDefault(s => s.IsSelected) ?? GetVisibleSheets().FirstOrDefault() ?? AllSheets.FirstOrDefault();
            var template = string.IsNullOrWhiteSpace(FileNameTemplate) ? "{SheetNumber} - {SheetName}" : FileNameTemplate;

            var value = template
                .Replace("{SheetNumber}", sample?.SheetNumber ?? "A101")
                .Replace("{SheetName}", sample?.SheetName ?? "Floor Plan")
                .Replace("{PaperSize}", sample?.PaperSize ?? "A1")
                .Replace("{ProjectNumber}", _projectNumber)
                .Replace("{ProjectName}", _projectName);

            foreach (var c in Path.GetInvalidFileNameChars())
                value = value.Replace(c, '-');

            value = value.Trim();
            if (string.IsNullOrWhiteSpace(value))
                value = "A101";

            return value + ".dwg";
        }

        private string GetProjectInfoValue(string propertyName, string fallback)
        {
            try
            {
                var projectInfo = _document?.ProjectInformation;
                if (projectInfo == null)
                    return fallback;

                var property = projectInfo.GetType().GetProperty(propertyName);
                var value = property == null ? null : property.GetValue(projectInfo, null);
                var text = value == null ? "" : value.ToString();
                return string.IsNullOrWhiteSpace(text) ? fallback : text;
            }
            catch
            {
                return fallback;
            }
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

        private void RefreshFilterState()
        {
            OnPropertyChanged(nameof(FilteredSheets));
            OnPropertyChanged(nameof(VisibleCount));
            OnPropertyChanged(nameof(TotalCount));
            OnPropertyChanged(nameof(SelectionSummary));
            OnPropertyChanged(nameof(AreVisibleSheetsSelected));
            OnPropertyChanged(nameof(SelectedSetsDisplay));
            OnPropertyChanged(nameof(FileNamePreview));
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void RefreshDerivedState()
        {
            UpdateSelectedCount();
            OnPropertyChanged(nameof(VisibleCount));
            OnPropertyChanged(nameof(TotalCount));
            OnPropertyChanged(nameof(SelectionSummary));
            OnPropertyChanged(nameof(AreVisibleSheetsSelected));
            OnPropertyChanged(nameof(ExportSummary));
            OnPropertyChanged(nameof(CombinedFilePreview));
            OnPropertyChanged(nameof(FileNamePreview));
            OnPropertyChanged(nameof(AutoCADStatusText));
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ReorderSelectedSheetsCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (SaveSheetSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private static Window GetActiveWindow()
        {
            return Application.Current?.Windows
                .OfType<Window>()
                .FirstOrDefault(w => w.IsActive);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private class InitialLoadData
        {
            public List<SheetInfo> Sheets { get; set; }
            public List<SheetScheduleInfo> Schedules { get; set; }
            public List<ViewSheetSetInfo> ViewSheetSets { get; set; }
            public List<string> ExportSetups { get; set; }
            public string ProjectName { get; set; }
            public string ProjectNumber { get; set; }
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
        private static readonly Regex NumberRegex = new Regex(@"(\d+)", RegexOptions.Compiled);

        public int Compare(string x, string y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            var xParts = NumberRegex.Split(x);
            var yParts = NumberRegex.Split(y);

            for (var i = 0; i < Math.Min(xParts.Length, yParts.Length); i++)
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
