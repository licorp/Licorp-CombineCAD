using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using Licorp_CombineCAD.Views;

namespace Licorp_CombineCAD.ViewModels
{
    public partial class ExportDialogViewModel
    {
        private void EditFileNameTemplate()
        {
            var dialog = new CustomFileNameDialog(FileNameTemplate, _projectNumber, _projectName)
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() == true)
                FileNameTemplate = dialog.FileNameTemplate;
        }

        private string BuildFileNamePreview()
        {
            var sample = AllSheets.FirstOrDefault(s => s.IsSelected) ?? GetVisibleSheets().FirstOrDefault() ?? AllSheets.FirstOrDefault();
            var template = string.IsNullOrWhiteSpace(FileNameTemplate) ? "{SheetNumber} - {SheetName}" : FileNameTemplate;

            var value = template
                .Replace("{SheetNumber}", sample?.SheetNumber ?? "A101")
                .Replace("{SheetName}", sample?.SheetName ?? "Floor Plan")
                .Replace("{PaperSize}", sample?.PaperSize ?? "A1")
                .Replace("{ProjectNumber}", _projectNumber)
                .Replace("{ProjectName}", _projectName);

            foreach (var c in Path.GetInvalidFileNameChars())
                value = value.Replace(c, '-');

            value = value.Trim();
            if (string.IsNullOrWhiteSpace(value))
                value = "A101";

            return value + ".dwg";
        }

        private string GetProjectInfoValue(string propertyName, string fallback)
        {
            try
            {
                var projectInfo = _document?.ProjectInformation;
                if (projectInfo == null)
                    return fallback;

                var property = projectInfo.GetType().GetProperty(propertyName);
                var value = property == null ? null : property.GetValue(projectInfo, null);
                var text = value == null ? "" : value.ToString();
                return string.IsNullOrWhiteSpace(text) ? fallback : text;
            }
            catch
            {
                return fallback;
            }
        }

        private string GetResolvedOutputFolder()
        {
            if (string.IsNullOrWhiteSpace(OutputFolder))
                return "";

            var projectFolderName = BuildProjectFolderName();
            return string.IsNullOrWhiteSpace(projectFolderName)
                ? OutputFolder
                : Path.Combine(OutputFolder, projectFolderName);
        }

        private string BuildProjectFileBaseName()
        {
            return BuildProjectFolderName();
        }

        private string BuildProjectFolderName()
        {
            var rawName = string.IsNullOrWhiteSpace(_projectName) ? "Project" : _projectName.Trim();
            foreach (var c in Path.GetInvalidFileNameChars())
                rawName = rawName.Replace(c, '-');

            rawName = rawName.Trim().Trim('.');
            return string.IsNullOrWhiteSpace(rawName) ? "Project" : rawName;
        }

        private void RefreshFilterState()
        {
            OnPropertyChanged(nameof(FilteredSheets));
            OnPropertyChanged(nameof(VisibleCount));
            OnPropertyChanged(nameof(TotalCount));
            OnPropertyChanged(nameof(SelectionSummary));
            OnPropertyChanged(nameof(AreVisibleSheetsSelected));
            OnPropertyChanged(nameof(SelectedSetsDisplay));
            OnPropertyChanged(nameof(FileNamePreview));
            OnPropertyChanged(nameof(ResolvedOutputFolder));
            OnPropertyChanged(nameof(OutputFolderHint));
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private void RefreshDerivedState()
        {
            UpdateSelectedCount();
            OnPropertyChanged(nameof(VisibleCount));
            OnPropertyChanged(nameof(TotalCount));
            OnPropertyChanged(nameof(SelectionSummary));
            OnPropertyChanged(nameof(AreVisibleSheetsSelected));
            OnPropertyChanged(nameof(ExportSummary));
            OnPropertyChanged(nameof(CombinedFilePreview));
            OnPropertyChanged(nameof(FileNamePreview));
            OnPropertyChanged(nameof(ResolvedOutputFolder));
            OnPropertyChanged(nameof(OutputFolderHint));
            OnPropertyChanged(nameof(AppVersionText));
            OnPropertyChanged(nameof(AutoCADStatusText));
            (ExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ReorderSelectedSheetsCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (SaveSheetSetCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }

        private static Window GetActiveWindow()
        {
            return Application.Current?.Windows
                .OfType<Window>()
                .FirstOrDefault(w => w.IsActive);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
