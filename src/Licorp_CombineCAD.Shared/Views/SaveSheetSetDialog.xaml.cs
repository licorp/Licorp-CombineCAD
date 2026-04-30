using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace Licorp_CombineCAD.Views
{
    public partial class SaveSheetSetDialog : Window, INotifyPropertyChanged
    {
        private string _setName;
        private string _errorMessage;

        public SaveSheetSetDialog(int selectedCount)
        {
            InitializeComponent();
            SelectedCount = selectedCount;
            SetName = "CombineCAD Set";
            DataContext = this;

            Loaded += (s, e) =>
            {
                SetNameTextBox.Focus();
                SetNameTextBox.SelectAll();
            };
        }

        public int SelectedCount { get; }

        public string InfoText => "This Sheet Set will contain " + SelectedCount + " selected sheet(s).";

        public string SetName
        {
            get => _setName;
            set
            {
                _setName = value;
                ErrorMessage = "";
                OnPropertyChanged();
            }
        }

        public string ErrorMessage
        {
            get => _errorMessage;
            set { _errorMessage = value; OnPropertyChanged(); }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            var value = (SetName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                ErrorMessage = "Enter a Sheet Set name.";
                return;
            }

            if (value.Equals("All Sheets", StringComparison.OrdinalIgnoreCase))
            {
                ErrorMessage = "This name is reserved. Choose a different name.";
                return;
            }

            var invalid = System.IO.Path.GetInvalidFileNameChars();
            if (value.Any(c => invalid.Contains(c)))
            {
                ErrorMessage = "Sheet Set name contains invalid characters.";
                return;
            }

            SetName = value;
            DialogResult = true;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
