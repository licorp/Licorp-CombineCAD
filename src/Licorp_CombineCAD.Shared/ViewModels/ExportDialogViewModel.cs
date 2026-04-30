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
using System.Reflection;

namespace Licorp_CombineCAD.ViewModels
{
    public partial class ExportDialogViewModel : INotifyPropertyChanged
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
        private string _statusMessage = "";
        private string _projectName = "Project";
        private string _projectNumber = "PRJ-001";
        private ExportMode _exportMode = ExportMode.MultiLayout;
        private bool _smartViewScale;
        private bool _openAfterExport;
        private bool _progressAlwaysOnTop = true;
        private bool _preserveCoincidentLines;
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
                OnPropertyChanged();
                ValidateExportMode();
                RefreshDerivedState();
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
            ? (string.IsNullOrWhiteSpace(AutoCADVersion) ? "AutoCAD detected" : string.Format("AutoCAD detected: {0}", AutoCADVersion))
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
            : Path.Combine(ResolvedOutputFolder, BuildProjectFileBaseName() + ".dwg");

        public string ResolvedOutputFolder => GetResolvedOutputFolder();
        public string OutputFolderHint => string.IsNullOrWhiteSpace(OutputFolder)
            ? "Choose a base folder. CombineCAD will create a project folder inside it."
            : string.Format("Base folder: {0}", OutputFolder);
        public string AppVersionText => "Licorp Ver " + (Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0");

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

        public void Dispose()
        {
            _isClosing = true;
            try { _cancellationTokenSource?.Cancel(); } catch { }
            try { _revitThreadService?.Dispose(); } catch { }
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
