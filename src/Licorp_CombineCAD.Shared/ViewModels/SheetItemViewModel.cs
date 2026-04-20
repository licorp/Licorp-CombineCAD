using System.ComponentModel;
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

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}