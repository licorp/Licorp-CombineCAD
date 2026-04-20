using System.Collections.Generic;
using System.ComponentModel;
using Autodesk.Revit.DB;

namespace Licorp_CombineCAD.Models
{
    /// <summary>
    /// Lightweight model for a Revit sheet, used for display and export.
    /// Refactored from Export+ SheetItem — removed legacy aliases and unused fields.
    /// </summary>
    public class SheetInfo : INotifyPropertyChanged
    {
        private bool _isSelected;
        private string _sheetNumber;
        private string _sheetName;
        private string _revision;
        private string _paperSize;
        private string _scaleText;
        private ElementId _elementId;
        private bool _hasNoView;

        /// <summary>
        /// Whether this sheet is selected for export
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected != value) { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }
        }

        /// <summary>
        /// Sheet Number (e.g. "A101")
        /// </summary>
        public string SheetNumber
        {
            get => _sheetNumber;
            set { if (_sheetNumber != value) { _sheetNumber = value; OnPropertyChanged(nameof(SheetNumber)); } }
        }

        /// <summary>
        /// Sheet Name (e.g. "Floor Plan - Level 1")
        /// </summary>
        public string SheetName
        {
            get => _sheetName;
            set { if (_sheetName != value) { _sheetName = value; OnPropertyChanged(nameof(SheetName)); } }
        }

        /// <summary>
        /// Current revision string
        /// </summary>
        public string Revision
        {
            get => _revision;
            set { if (_revision != value) { _revision = value; OnPropertyChanged(nameof(Revision)); } }
        }

        /// <summary>
        /// Detected paper size (e.g. "A3", "A1")
        /// </summary>
        public string PaperSize
        {
            get => _paperSize;
            set { if (_paperSize != value) { _paperSize = value; OnPropertyChanged(nameof(PaperSize)); } }
        }

        /// <summary>
        /// Display scale text (e.g. "1:100" or "As Indicated")
        /// </summary>
        public string ScaleText
        {
            get => _scaleText;
            set { if (_scaleText != value) { _scaleText = value; OnPropertyChanged(nameof(ScaleText)); } }
        }

        /// <summary>
        /// Revit ElementId of the ViewSheet
        /// </summary>
        public ElementId ElementId
        {
            get => _elementId;
            set { if (_elementId != value) { _elementId = value; OnPropertyChanged(nameof(ElementId)); } }
        }

        /// <summary>
        /// True if no viewports are placed on this sheet
        /// </summary>
        public bool HasNoView
        {
            get => _hasNoView;
            set { if (_hasNoView != value) { _hasNoView = value; OnPropertyChanged(nameof(HasNoView)); } }
        }

        /// <summary>
        /// Primary view scale (integer, e.g. 100 for 1:100). 0 if multiple or unknown.
        /// </summary>
        public int PrimaryScale { get; set; }

        /// <summary>
        /// All view scales on this sheet
        /// </summary>
        public List<int> ViewScales { get; set; } = new List<int>();

        // ===== INotifyPropertyChanged =====
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
