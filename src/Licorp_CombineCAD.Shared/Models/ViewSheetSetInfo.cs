using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Licorp_CombineCAD.Models
{
    public class ViewSheetSetInfo : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string Name { get; set; }
        public bool IsBuiltIn { get; set; }
        public List<string> SheetIdValues { get; } = new List<string>();

        public string DisplayName
        {
            get
            {
                var count = SheetIdValues == null ? 0 : SheetIdValues.Count;
                return string.Format("{0} ({1})", Name, count);
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                    return;

                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public override string ToString()
        {
            return DisplayName;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
