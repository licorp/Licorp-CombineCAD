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
        private string _paperSize;
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
        /// Detected paper size (e.g. "A3", "A1")
        /// </summary>
        public string PaperSize
        {
            get => _paperSize;
            set { if (_paperSize != value) { _paperSize = value; OnPropertyChanged(nameof(PaperSize)); } }
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

        // ===== INotifyPropertyChanged =====
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
