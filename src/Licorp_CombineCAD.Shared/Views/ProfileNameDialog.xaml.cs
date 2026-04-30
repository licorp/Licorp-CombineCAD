using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace Licorp_CombineCAD.Views
{
    public partial class ProfileNameDialog : Window, INotifyPropertyChanged
    {
        private string _profileName;
        private string _errorMessage;

        public ProfileNameDialog(string prompt, string initialName = "")
        {
            InitializeComponent();
            Prompt = prompt;
            ProfileName = initialName ?? "";
            DataContext = this;
            Loaded += (s, e) =>
            {
                NameTextBox.Focus();
                NameTextBox.SelectAll();
            };
        }

        public string Prompt { get; }

        public string ProfileName
        {
            get => _profileName;
            set
            {
                _profileName = value;
                ErrorMessage = "";
                OnPropertyChanged();
            }
        }

        public string ErrorMessage
        {
            get => _errorMessage;
            set { _errorMessage = value; OnPropertyChanged(); }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var value = (ProfileName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                ErrorMessage = "Enter a profile name.";
                return;
            }

            var invalid = System.IO.Path.GetInvalidFileNameChars();
            if (value.Any(c => invalid.Contains(c)))
            {
                ErrorMessage = "Profile name contains invalid file-name characters.";
                return;
            }

            ProfileName = value;
            DialogResult = true;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
