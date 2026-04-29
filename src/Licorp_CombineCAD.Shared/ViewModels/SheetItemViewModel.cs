using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.ViewModels
{
    /// <summary>
    /// ViewModel for a single sheet item in the export dialog.
    /// Wraps SheetInfo with UI-friendly properties.
    /// </summary>
    public class SheetItemViewModel : INotifyPropertyChanged
    {
        private readonly SheetInfo _model;
        private bool _isSelected;

        public SheetItemViewModel(SheetInfo model)
        {
            _model = model ?? throw new System.ArgumentNullException(nameof(model));
            _isSelected = model.IsSelected;
        }

        public SheetInfo Model => _model;

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    _model.IsSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public string SheetNumber => _model.SheetNumber;
        public string SheetName => _model.SheetName;
        public string Revision => _model.Revision;
        public string PaperSize => _model.PaperSize;
        public string ScaleText => _model.ScaleText;
        public ElementId ElementId => _model.ElementId;
        public bool HasNoView => _model.HasNoView;

        public string DisplayText => $"{SheetNumber} - {SheetName}";

        public string StatusText
        {
            get
            {
                if (HasNoView)
                    return "No views";
                if (_model.ViewScales != null && _model.ViewScales.Distinct().Count() > 1)
                    return "Mixed scale";
                return "OK";
            }
        }

        public string StatusToolTip
        {
            get
            {
                if (HasNoView)
                    return "No model viewport detected. This sheet can still export title blocks, schedules, images, and annotations.";
                if (_model.ViewScales != null && _model.ViewScales.Distinct().Count() > 1)
                    return "Multiple view scales detected on this sheet.";
                return "Ready for export.";
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
