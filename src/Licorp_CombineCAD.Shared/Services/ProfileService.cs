using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using Newtonsoft.Json;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Profile service — saves and loads last-used export settings.
    /// JSON-based config stored in %APPDATA%\Licorp\CombiCAD\
    /// </summary>
    public class ProfileService
    {
        private readonly string _configFolder;
        private readonly string _configFile;
        private ExportProfile _lastProfile;

        public ProfileService()
        {
            _configFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Licorp", "CombiCAD");
            _configFile = Path.Combine(_configFolder, "LastProfile.json");

            EnsureFolderExists();
        }

        public ExportProfile LoadLastProfile()
        {
            try
            {
                if (File.Exists(_configFile))
                {
                    var json = File.ReadAllText(_configFile);
                    _lastProfile = JsonConvert.DeserializeObject<ExportProfile>(json);
                    System.Diagnostics.Debug.WriteLine($"[Profile] Loaded: {_configFile}");
                    return _lastProfile;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Profile] Load error: {ex.Message}");
            }

            _lastProfile = new ExportProfile();
            return _lastProfile;
        }

        public void SaveLastProfile(ExportProfile profile)
        {
            try
            {
                var json = JsonConvert.SerializeObject(profile, Formatting.Indented);
                File.WriteAllText(_configFile, json);
                _lastProfile = profile;
                System.Diagnostics.Debug.WriteLine($"[Profile] Saved: {_configFile}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Profile] Save error: {ex.Message}");
            }
        }

        private void EnsureFolderExists()
        {
            if (!Directory.Exists(_configFolder))
            {
                Directory.CreateDirectory(_configFolder);
            }
        }
    }

    public class ExportProfile
    {
        public string OutputFolder { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public string FileNameTemplate { get; set; } = "{SheetNumber} - {SheetName}";
        public string SelectedSetup { get; set; } = "";
public string ExportMode { get; set; } = "MultiLayout";
    public string DwgVersion { get; set; } = "2018";
    public bool AutoBindXRef { get; set; } = true;
    public bool SmartViewScale { get; set; } = false;
    public bool OpenAfterExport { get; set; } = false;
    public bool PreserveCoincidentLines { get; set; } = false;
    public bool CreateSubfolders { get; set; } = false;
        public bool HideScopeBoxes { get; set; } = true;
        public bool HideReferencePlanes { get; set; } = true;
        public DateTime LastUsed { get; set; } = DateTime.Now;
    }
}