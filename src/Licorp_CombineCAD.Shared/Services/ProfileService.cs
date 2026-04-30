using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace Licorp_CombineCAD.Services
{
    public class ProfileService
    {
        private const string DefaultProfileName = "Default";

        private readonly string _configFolder;
        private readonly string _profilesFolder;
        private readonly string _legacyConfigFile;
        private readonly string _lastProfileIdFile;

        public ProfileService()
        {
            _configFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Licorp", "CombiCAD");
            _profilesFolder = Path.Combine(_configFolder, "Profiles");
            _legacyConfigFile = Path.Combine(_configFolder, "LastProfile.json");
            _lastProfileIdFile = Path.Combine(_configFolder, "LastProfileId.txt");

            Directory.CreateDirectory(_configFolder);
            Directory.CreateDirectory(_profilesFolder);
        }

        public List<ExportProfile> LoadProfiles()
        {
            MigrateLegacyProfileIfNeeded();

            var profiles = new List<ExportProfile>();
            foreach (var file in Directory.GetFiles(_profilesFolder, "*.json"))
            {
                try
                {
                    var profile = JsonConvert.DeserializeObject<ExportProfile>(File.ReadAllText(file));
                    if (profile == null)
                        continue;

                    NormalizeProfile(profile);
                    profiles.Add(profile);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine("[Profile] Failed to read profile '" + file + "': " + ex.Message);
                }
            }

            if (profiles.Count == 0)
            {
                var defaultProfile = CreateDefaultProfile();
                SaveProfile(defaultProfile);
                SetLastProfile(defaultProfile.Id);
                profiles.Add(defaultProfile);
            }

            return profiles
                .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public ExportProfile LoadLastProfile()
        {
            var profiles = LoadProfiles();
            var lastId = ReadLastProfileId();

            var profile = profiles.FirstOrDefault(p => string.Equals(p.Id, lastId, StringComparison.OrdinalIgnoreCase))
                ?? profiles.FirstOrDefault(p => string.Equals(p.Name, DefaultProfileName, StringComparison.OrdinalIgnoreCase))
                ?? profiles.FirstOrDefault();

            return profile ?? CreateDefaultProfile();
        }

        public void SaveProfile(ExportProfile profile)
        {
            if (profile == null)
                return;

            try
            {
                NormalizeProfile(profile);
                profile.LastModified = DateTime.Now;
                profile.LastUsed = DateTime.Now;

                var json = JsonConvert.SerializeObject(profile, Formatting.Indented);
                File.WriteAllText(GetProfilePath(profile.Id), json);
                Trace.WriteLine("[Profile] Saved profile: " + profile.Name);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Profile] Save error: " + ex.Message);
            }
        }

        public void SaveLastProfile(ExportProfile profile)
        {
            SaveProfile(profile);
            if (profile == null)
                return;

            SetLastProfile(profile.Id);

            try
            {
                var json = JsonConvert.SerializeObject(profile, Formatting.Indented);
                File.WriteAllText(_legacyConfigFile, json);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Profile] Legacy mirror save error: " + ex.Message);
            }
        }

        public void DeleteProfile(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return;

            try
            {
                var path = GetProfilePath(id);
                if (File.Exists(path))
                    File.Delete(path);

                if (string.Equals(ReadLastProfileId(), id, StringComparison.OrdinalIgnoreCase))
                    File.Delete(_lastProfileIdFile);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Profile] Delete error: " + ex.Message);
            }
        }

        public void SetLastProfile(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return;

            try
            {
                File.WriteAllText(_lastProfileIdFile, id);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Profile] Last profile save error: " + ex.Message);
            }
        }

        private void MigrateLegacyProfileIfNeeded()
        {
            try
            {
                if (Directory.GetFiles(_profilesFolder, "*.json").Length > 0 || !File.Exists(_legacyConfigFile))
                    return;

                var legacy = JsonConvert.DeserializeObject<ExportProfile>(File.ReadAllText(_legacyConfigFile))
                    ?? CreateDefaultProfile();
                legacy.Id = string.IsNullOrWhiteSpace(legacy.Id) ? Guid.NewGuid().ToString("N") : legacy.Id;
                legacy.Name = string.IsNullOrWhiteSpace(legacy.Name) ? DefaultProfileName : legacy.Name;
                legacy.Description = string.IsNullOrWhiteSpace(legacy.Description)
                    ? "Migrated from LastProfile.json"
                    : legacy.Description;
                SaveProfile(legacy);
                SetLastProfile(legacy.Id);
                Trace.WriteLine("[Profile] Migrated legacy LastProfile.json");
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Profile] Migration error: " + ex.Message);
            }
        }

        private ExportProfile CreateDefaultProfile()
        {
            return new ExportProfile
            {
                Id = Guid.NewGuid().ToString("N"),
                Name = DefaultProfileName,
                Description = "Default CombineCAD profile",
                OutputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                FileNameTemplate = "{SheetNumber} - {SheetName}",
                ExportMode = "MultiLayout",
                DwgVersion = "2018",
                AutoBindXRef = true,
                MergeEngine = "AcCoreConsole",
                VerifyCombinedDwg = true,
                CreateSheetSet = true
            };
        }

        private void NormalizeProfile(ExportProfile profile)
        {
            if (string.IsNullOrWhiteSpace(profile.Id))
                profile.Id = Guid.NewGuid().ToString("N");
            if (string.IsNullOrWhiteSpace(profile.Name))
                profile.Name = DefaultProfileName;
            if (string.IsNullOrWhiteSpace(profile.OutputFolder))
                profile.OutputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (string.IsNullOrWhiteSpace(profile.FileNameTemplate))
                profile.FileNameTemplate = "{SheetNumber} - {SheetName}";
            if (string.IsNullOrWhiteSpace(profile.ExportMode))
                profile.ExportMode = "MultiLayout";
            if (string.IsNullOrWhiteSpace(profile.DwgVersion))
                profile.DwgVersion = "2018";
            if (string.IsNullOrWhiteSpace(profile.SortMode))
                profile.SortMode = "Sheet Number";
            if (string.IsNullOrWhiteSpace(profile.VerticalAlign))
                profile.VerticalAlign = "Top";
            if (string.IsNullOrWhiteSpace(profile.RasterImageMode))
                profile.RasterImageMode = "KeepReference";
            if (string.IsNullOrWhiteSpace(profile.MergeEngine))
                profile.MergeEngine = "AcCoreConsole";
            if (profile.SchemaVersion <= 0)
                profile.SchemaVersion = 2;
        }

        private string ReadLastProfileId()
        {
            try
            {
                return File.Exists(_lastProfileIdFile)
                    ? File.ReadAllText(_lastProfileIdFile).Trim()
                    : "";
            }
            catch
            {
                return "";
            }
        }

        private string GetProfilePath(string id)
        {
            return Path.Combine(_profilesFolder, id + ".json");
        }
    }

    public class ExportProfile
    {
        public int SchemaVersion { get; set; } = 2;
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Name { get; set; } = "Default";
        public string Description { get; set; } = "";
        public DateTime LastModified { get; set; } = DateTime.Now;

        public string OutputFolder { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public string FileNameTemplate { get; set; } = "{SheetNumber} - {SheetName}";
        public string SelectedSetup { get; set; } = "";
        public string ExportMode { get; set; } = "MultiLayout";
        public string DwgVersion { get; set; } = "2018";
        public bool AutoBindXRef { get; set; } = true;
        public bool SmartViewScale { get; set; } = false;
        public bool OpenAfterExport { get; set; } = false;
        public bool ProgressAlwaysOnTop { get; set; } = true;
        public bool PreserveCoincidentLines { get; set; } = false;
        public bool CreateSubfolders { get; set; } = false;
        public string SortMode { get; set; } = "Sheet Number";
        public string SelectedSheetScheduleId { get; set; } = "";
        public string VerticalAlign { get; set; } = "Top";
        public bool CreateSheetSet { get; set; } = true;
        public string RasterImageMode { get; set; } = "KeepReference";
        public string MergeEngine { get; set; } = "AcCoreConsole";
        public bool VerifyCombinedDwg { get; set; } = true;
        public bool HideScopeBoxes { get; set; } = true;
        public bool HideReferencePlanes { get; set; } = true;
        public DateTime LastUsed { get; set; } = DateTime.Now;
    }
}
