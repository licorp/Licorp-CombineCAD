using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Licorp_CombineCAD.Models
{
    /// <summary>
    /// All DWG export settings configurable by the user.
    /// Refactored from Export+ DWGExportSettings + ExportSettings.
    /// </summary>
    public class ExportSettings : INotifyPropertyChanged
    {
        // ===== Output =====
        private string _outputFolder = "";
        private string _fileNameTemplate = "{SheetNumber} - {SheetName}";
        private ExportMode _exportMode = ExportMode.MultiLayout;

        // ===== DWG Options =====
        private string _dwgExportSetupName = "";
        private string _dwgVersion = "2018";

        // ===== Advanced Options (MLabs features) =====
        private bool _openAfterExport = false;
        private bool _smartViewScale = false;
        private bool _progressAlwaysOnTop = true;
        private string _orderRuleSource = "Sheet Number";
        private string _selectedSheetScheduleId = "";

        // ===== Geometry Options =====
        private bool _preserveCoincidentLines = false;
        private bool _mergeLayers = true;

        private VerticalAlignment _verticalAlign = VerticalAlignment.Top;
        private SortMode _sortMode = SortMode.SheetNumber;
        private string _modelSpaceArrangement = "Horizontal";
        private int _gridColumns = 3;
        private double _customSpacing = 50.0;
        private bool _reverseSortOrder = false;

        // ===== Layout Name Template (Phase 3) =====
        private string _layoutNameTemplate = "{SheetNumber} - {SheetName}";

        // ===== Output =====
        public string OutputFolder
        {
            get => _outputFolder;
            set { _outputFolder = value; OnPropertyChanged(); }
        }

        public string FileNameTemplate
        {
            get => _fileNameTemplate;
            set { _fileNameTemplate = value; OnPropertyChanged(); }
        }

        public ExportMode ExportMode
        {
            get => _exportMode;
            set { _exportMode = value; OnPropertyChanged(); }
        }

        // ===== DWG Options =====
        public string DwgExportSetupName
        {
            get => _dwgExportSetupName;
            set { _dwgExportSetupName = value; OnPropertyChanged(); }
        }

        public string DwgVersion
        {
            get => _dwgVersion;
            set { _dwgVersion = value; OnPropertyChanged(); }
        }

        // ===== Advanced Options =====
        public bool OpenAfterExport
        {
            get => _openAfterExport;
            set { _openAfterExport = value; OnPropertyChanged(); }
        }

        public bool SmartViewScale
        {
            get => _smartViewScale;
            set { _smartViewScale = value; OnPropertyChanged(); }
        }

        public bool ProgressAlwaysOnTop
        {
            get => _progressAlwaysOnTop;
            set { _progressAlwaysOnTop = value; OnPropertyChanged(); }
        }

        public string OrderRuleSource
        {
            get => _orderRuleSource;
            set { _orderRuleSource = value; OnPropertyChanged(); }
        }

        public string SelectedSheetScheduleId
        {
            get => _selectedSheetScheduleId;
            set { _selectedSheetScheduleId = value; OnPropertyChanged(); }
        }

        // ===== Geometry Options =====
        public bool PreserveCoincidentLines
        {
            get => _preserveCoincidentLines;
            set { _preserveCoincidentLines = value; OnPropertyChanged(); }
        }

        public bool MergeLayers
        {
            get => _mergeLayers;
            set { _mergeLayers = value; OnPropertyChanged(); }
        }

        public VerticalAlignment VerticalAlign
        {
            get => _verticalAlign;
            set { _verticalAlign = value; OnPropertyChanged(); }
        }

        public SortMode SortMode
        {
            get => _sortMode;
            set { _sortMode = value; OnPropertyChanged(); }
        }

        public string ModelSpaceArrangement
        {
            get => _modelSpaceArrangement;
            set { _modelSpaceArrangement = value; OnPropertyChanged(); }
        }

        public int GridColumns
        {
            get => _gridColumns;
            set { _gridColumns = value; OnPropertyChanged(); }
        }

        public double CustomSpacing
        {
            get => _customSpacing;
            set { _customSpacing = value; OnPropertyChanged(); }
        }

        public bool ReverseSortOrder
        {
            get => _reverseSortOrder;
            set { _reverseSortOrder = value; OnPropertyChanged(); }
        }

        // ===== Layout Name Template (Phase 3) =====
        /// <summary>
        /// Template for layout names in merged DWG.
        /// Available placeholders: {SheetNumber}, {SheetName}, {PaperSize}
        /// Example: "{PaperSize} - {SheetName}" produces "A1 - Floor Plan"
        /// </summary>
        public string LayoutNameTemplate
        {
            get => _layoutNameTemplate;
            set { _layoutNameTemplate = value; OnPropertyChanged(); }
        }

        // ===== INotifyPropertyChanged =====
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
