using System;
using System.Linq;
using System.Windows;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.Services;
using Licorp_CombineCAD.Views;

namespace Licorp_CombineCAD.ViewModels
{
    public partial class ExportDialogViewModel
    {
        private void ApplySelectedProfile()
        {
            if (SelectedProfile == null)
                return;

            _profile = SelectedProfile;
            _profileService.SetLastProfile(_profile.Id);
            ApplyProfileValues(_profile);
            StatusMessage = "Profile applied: " + _profile.Name;
        }

        private void CreateProfileFromCurrentSettings()
        {
            var dialog = new ProfileNameDialog("New profile name")
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() != true)
                return;

            if (Profiles.Any(p => string.Equals(p.Name, dialog.ProfileName, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Profile already exists: " + dialog.ProfileName, "Profile", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var profile = new ExportProfile
            {
                Id = Guid.NewGuid().ToString("N"),
                Name = dialog.ProfileName,
                Description = "Created " + DateTime.Now.ToString("yyyy-MM-dd HH:mm")
            };

            SaveCurrentSettingsToProfile(profile);
            _profileService.SaveLastProfile(profile);
            Profiles.Add(profile);
            SelectedProfile = profile;
            _profile = profile;
            StatusMessage = "Profile created: " + profile.Name;
        }

        private void DeleteSelectedProfile()
        {
            if (SelectedProfile == null || Profiles.Count <= 1)
                return;

            var confirm = MessageBox.Show(
                "Delete profile '" + SelectedProfile.Name + "'?",
                "Delete Profile",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes)
                return;

            var deleted = SelectedProfile;
            _profileService.DeleteProfile(deleted.Id);
            Profiles.Remove(deleted);
            SelectedProfile = Profiles.FirstOrDefault();
            _profile = SelectedProfile;

            if (_profile != null)
            {
                _profileService.SetLastProfile(_profile.Id);
                ApplyProfileValues(_profile);
            }
        }

        private void SaveSettings()
        {
            if (SelectedProfile == null)
            {
                CreateProfileFromCurrentSettings();
                return;
            }

            _profile = SelectedProfile;
            SaveCurrentSettingsToProfile(_profile);
            _profileService.SaveLastProfile(_profile);
            StatusMessage = "Profile saved: " + _profile.Name;
        }

        private void SaveCurrentSettingsToProfile(ExportProfile profile)
        {
            profile.OutputFolder = OutputFolder;
            profile.SelectedSetup = SelectedSetup;
            profile.ExportMode = ExportMode.ToString();
            profile.DwgVersion = SelectedDwgVersion;
            profile.FileNameTemplate = FileNameTemplate;
            profile.SmartViewScale = SmartViewScale;
            profile.OpenAfterExport = OpenAfterExport;
            profile.ProgressAlwaysOnTop = ProgressAlwaysOnTop;
            profile.PreserveCoincidentLines = PreserveCoincidentLines;
            profile.SortMode = SelectedSortMode;
            profile.SelectedSheetScheduleId = SelectedSheetSchedule?.ElementIdValue ?? "";
            profile.VerticalAlign = SelectedVerticalAlignment;
            profile.LastUsed = DateTime.Now;
        }

        private void ApplyProfileValues(ExportProfile profile)
        {
            if (profile == null)
                return;

            OutputFolder = profile.OutputFolder;
            SelectedSetup = profile.SelectedSetup;

            if (Enum.TryParse(profile.ExportMode, out ExportMode mode))
                ExportMode = mode;

            if (!string.IsNullOrEmpty(profile.DwgVersion))
                SelectedDwgVersion = profile.DwgVersion;

            if (!string.IsNullOrEmpty(profile.FileNameTemplate))
                FileNameTemplate = profile.FileNameTemplate;

            SmartViewScale = profile.SmartViewScale;
            OpenAfterExport = profile.OpenAfterExport;
            ProgressAlwaysOnTop = profile.ProgressAlwaysOnTop;
            PreserveCoincidentLines = profile.PreserveCoincidentLines;

            if (!string.IsNullOrEmpty(profile.SortMode) && AvailableSortModes.Contains(profile.SortMode))
                SelectedSortMode = profile.SortMode;

            if (!string.IsNullOrEmpty(profile.VerticalAlign) && AvailableVerticalAlignments.Contains(profile.VerticalAlign))
                SelectedVerticalAlignment = profile.VerticalAlign;

            ApplySavedScheduleSelection();
            RefreshDerivedState();
        }
    }
}
