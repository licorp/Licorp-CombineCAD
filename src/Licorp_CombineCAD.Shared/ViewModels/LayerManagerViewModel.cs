using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Licorp_CombineCAD.Services;

namespace Licorp_CombineCAD.ViewModels
{
    /// <summary>
    /// ViewModel for the Layer Manager dialog.
    /// Handles export/import of DWG layer mapping settings.
    /// </summary>
    public class LayerManagerViewModel : INotifyPropertyChanged
    {
        private readonly UIDocument _uiDocument;
        private readonly Document _document;
        private readonly LayerMappingService _layerService;

        private string _selectedSetup;
        private string _statusMessage = "";
        private bool _isProcessing = false;

        public LayerManagerViewModel(UIDocument uiDocument)
        {
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _document = uiDocument.Document;
            _layerService = new LayerMappingService(_document);

            LoadExportSetups();

            ExportCommand = new RelayCommand(() => ExportMapping(), () => !IsProcessing && !string.IsNullOrEmpty(SelectedSetup));
            ImportCommand = new RelayCommand(() => ImportMapping(), () => !IsProcessing && !string.IsNullOrEmpty(SelectedSetup));
        }

        public ObservableCollection<string> AvailableSetups { get; } = new ObservableCollection<string>();

        public string SelectedSetup
        {
            get => _selectedSetup;
            set { _selectedSetup = value; OnPropertyChanged(); }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set
            {
                if (_isProcessing != value)
                {
                    _isProcessing = value;
                    OnPropertyChanged();
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public ICommand ExportCommand { get; }
        public ICommand ImportCommand { get; }

        private void LoadExportSetups()
        {
            try
            {
                var collector = new FilteredElementCollector(_document)
                    .OfClass(typeof(ExportDWGSettings));

                foreach (ExportDWGSettings setting in collector)
                {
                    if (!string.IsNullOrEmpty(setting.Name))
                        AvailableSetups.Add(setting.Name);
                }

                if (AvailableSetups.Count > 0)
                    SelectedSetup = AvailableSetups.First();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LayerManager] Error loading setups: {ex.Message}");
            }
        }

        private void ExportMapping()
        {
            if (string.IsNullOrEmpty(SelectedSetup)) return;

            var dialog = new System.Windows.Forms.SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                FileName = $"LayerMapping_{SelectedSetup}_{DateTime.Now:yyyyMMdd}.txt",
                Title = "Export Layer Mapping"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                IsProcessing = true;
                StatusMessage = "Exporting...";

                try
                {
                    var success = _layerService.ExportLayerMapping(SelectedSetup, dialog.FileName);
                    StatusMessage = success ? $"Exported to {dialog.FileName}" : "Export failed";
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error: {ex.Message}";
                }
                finally
                {
                    IsProcessing = false;
                }
            }
        }

        private void ImportMapping()
        {
            if (string.IsNullOrEmpty(SelectedSetup)) return;

            var dialog = new System.Windows.Forms.OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                Title = "Import Layer Mapping"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                IsProcessing = true;
                StatusMessage = "Importing...";

                try
                {
                    var success = _layerService.ImportLayerMapping(SelectedSetup, dialog.FileName);
                    StatusMessage = success ? "Import completed" : "Import failed";
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Error: {ex.Message}";
                }
                finally
                {
                    IsProcessing = false;
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}