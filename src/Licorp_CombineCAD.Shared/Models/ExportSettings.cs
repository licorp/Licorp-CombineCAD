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
        private ExportMode _exportMode = ExportMode.Individual;

        // ===== DWG Options =====
        private string _dwgExportSetupName = "";
        private string _dwgVersion = "2018";
        private bool _autoBindXRef = true;
        private bool _useSharedCoordinates = false;

        // ===== Advanced Options (MLabs features) =====
        private bool _openAfterExport = false;
        private bool _smartViewScale = false;
        private bool _progressAlwaysOnTop = true;

        // ===== Geometry Options =====
        private bool _hideScopeBoxes = true;
        private bool _hideReferencePlanes = true;
        private bool _hideUnreferencedViewTags = true;
        private bool _preserveCoincidentLines = false;

        private VerticalAlignment _verticalAlign = VerticalAlignment.Top;
        private SortMode _sortMode = SortMode.SheetNumber;

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

        public bool AutoBindXRef
        {
            get => _autoBindXRef;
            set { _autoBindXRef = value; OnPropertyChanged(); }
        }

        public bool UseSharedCoordinates
        {
            get => _useSharedCoordinates;
            set { _useSharedCoordinates = value; OnPropertyChanged(); }
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

        // ===== Geometry Options =====
        public bool HideScopeBoxes
        {
            get => _hideScopeBoxes;
            set { _hideScopeBoxes = value; OnPropertyChanged(); }
        }

        public bool HideReferencePlanes
        {
            get => _hideReferencePlanes;
            set { _hideReferencePlanes = value; OnPropertyChanged(); }
        }

        public bool HideUnreferencedViewTags
        {
            get => _hideUnreferencedViewTags;
            set { _hideUnreferencedViewTags = value; OnPropertyChanged(); }
        }

        public bool PreserveCoincidentLines
        {
            get => _preserveCoincidentLines;
            set { _preserveCoincidentLines = value; OnPropertyChanged(); }
        }

        private bool _createSubfolders = false;
        public bool CreateSubfolders
        {
            get => _createSubfolders;
            set { _createSubfolders = value; OnPropertyChanged(); }
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

        // ===== INotifyPropertyChanged =====
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
