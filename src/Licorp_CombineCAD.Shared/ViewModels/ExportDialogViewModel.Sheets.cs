using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.Services;
using Licorp_CombineCAD.Views;

namespace Licorp_CombineCAD.ViewModels
{
    public partial class ExportDialogViewModel
    {
        public async Task InitializeAsync()
        {
            if (_isInitialized)
                return;

            _isInitialized = true;
            IsLoadingSheets = true;
            LoadingMessage = "Loading sheet number, sheet name, and sheet sets...";

            try
            {
                var data = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    return new InitialLoadData
                    {
                        Sheets = _sheetCollector.GetAllSheets(),
                        Schedules = _sheetScheduleService.GetSheetListSchedules(),
                        ViewSheetSets = _viewSheetSetService.GetSheetSets(),
                        ExportSetups = _exportService.GetAvailableExportSetups(),
                        ProjectName = GetProjectInfoValue("Name", "Project"),
                        ProjectNumber = GetProjectInfoValue("Number", "PRJ-001")
                    };
                });

                if (_isClosing)
                    return;

                _projectName = data.ProjectName;
                _projectNumber = data.ProjectNumber;
                LoadExportSetups(data.ExportSetups);
                LoadSheetSchedules(data.Schedules);
                LoadViewSheetSets(data.ViewSheetSets);
                LoadSheets(data.Sheets);
                ApplySavedScheduleSelection();
                StatusMessage = "Ready.";
            }
            catch (Exception ex)
            {
                StatusMessage = "Failed to load sheets: " + ex.Message;
                Trace.WriteLine("[ExportDialog] Load failed: " + ex);
            }
            finally
            {
                IsLoadingSheets = false;
                LoadingMessage = "";
                RefreshDerivedState();
            }
        }

        public void ReorderSheet(SheetItemViewModel source, SheetItemViewModel target)
        {
            var sourceIndex = AllSheets.IndexOf(source);
            var targetIndex = AllSheets.IndexOf(target);
            if (sourceIndex < 0 || targetIndex < 0 || sourceIndex == targetIndex)
                return;

            AllSheets.Move(sourceIndex, targetIndex);
            SelectedSortMode = "Custom";
            RefreshFilterState();
        }

        private void LoadExportSetups(IEnumerable<string> setups)
        {
            AvailableSetups.Clear();
            foreach (var setup in setups ?? Enumerable.Empty<string>())
                AvailableSetups.Add(setup);

            if (string.IsNullOrWhiteSpace(SelectedSetup) && AvailableSetups.Count > 0)
                SelectedSetup = AvailableSetups[0];
        }

        private void LoadSheets(IEnumerable<SheetInfo> sheets)
        {
            foreach (var old in AllSheets)
                old.PropertyChanged -= OnSheetSelectionChanged;

            AllSheets.Clear();
            foreach (var sheet in sheets ?? Enumerable.Empty<SheetInfo>())
            {
                var vm = new SheetItemViewModel(sheet);
                vm.PropertyChanged += OnSheetSelectionChanged;
                AllSheets.Add(vm);
            }

            ApplySort();
            UpdateSelectedCount();
            RefreshFilterState();
        }

        private void LoadSheetSchedules(IEnumerable<SheetScheduleInfo> schedules)
        {
            AvailableSheetSchedules.Clear();
            foreach (var schedule in schedules ?? Enumerable.Empty<SheetScheduleInfo>())
                AvailableSheetSchedules.Add(schedule);

            if (SelectedSheetSchedule == null && AvailableSheetSchedules.Count > 0)
                SelectedSheetSchedule = AvailableSheetSchedules[0];
        }

        private void LoadViewSheetSets(IEnumerable<ViewSheetSetInfo> sets)
        {
            foreach (var old in AvailableViewSheetSets)
                old.PropertyChanged -= OnViewSheetSetChanged;

            AvailableViewSheetSets.Clear();
            foreach (var set in sets ?? Enumerable.Empty<ViewSheetSetInfo>())
            {
                set.PropertyChanged += OnViewSheetSetChanged;
                AvailableViewSheetSets.Add(set);
            }

            var allSheets = AvailableViewSheetSets.FirstOrDefault(s => s.IsBuiltIn);
            if (allSheets != null)
                allSheets.IsSelected = true;

            RefreshFilterState();
        }

        private void ApplySavedScheduleSelection()
        {
            if (string.IsNullOrWhiteSpace(_profile?.SelectedSheetScheduleId))
                return;

            var saved = AvailableSheetSchedules.FirstOrDefault(s => s.ElementIdValue == _profile.SelectedSheetScheduleId);
            if (saved != null)
                SelectedSheetSchedule = saved;
        }

        private List<SheetItemViewModel> GetVisibleSheets()
        {
            IEnumerable<SheetItemViewModel> query = AllSheets;
            var selectedSets = AvailableViewSheetSets.Where(s => s.IsSelected).ToList();
            var allSheetsSelected = selectedSets.Count == 0 || selectedSets.Any(s => s.IsBuiltIn);

            if (!allSheetsSelected)
            {
                var allowedIds = new HashSet<string>(
                    selectedSets.SelectMany(s => s.SheetIdValues),
                    StringComparer.OrdinalIgnoreCase);

                query = query.Where(s => allowedIds.Contains(ViewSheetSetService.GetElementIdValue(s.ElementId)));
            }

            if (!string.IsNullOrWhiteSpace(FilterText))
            {
                var filter = FilterText.Trim().ToLowerInvariant();
                query = query.Where(s =>
                    (s.SheetNumber ?? "").ToLowerInvariant().Contains(filter) ||
                    (s.SheetName ?? "").ToLowerInvariant().Contains(filter) ||
                    (s.PaperSize ?? "").ToLowerInvariant().Contains(filter));
            }

            return query.ToList();
        }

        private void OnViewSheetSetChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(ViewSheetSetInfo.IsSelected) || _isUpdatingSetSelection)
                return;

            try
            {
                _isUpdatingSetSelection = true;
                var changed = sender as ViewSheetSetInfo;
                if (changed != null && changed.IsSelected)
                {
                    if (changed.IsBuiltIn)
                    {
                        foreach (var set in AvailableViewSheetSets.Where(s => !s.IsBuiltIn))
                            set.IsSelected = false;
                    }
                    else
                    {
                        foreach (var set in AvailableViewSheetSets.Where(s => s.IsBuiltIn))
                            set.IsSelected = false;
                    }
                }

                if (!AvailableViewSheetSets.Any(s => s.IsSelected))
                {
                    var allSheets = AvailableViewSheetSets.FirstOrDefault(s => s.IsBuiltIn);
                    if (allSheets != null)
                        allSheets.IsSelected = true;
                }
            }
            finally
            {
                _isUpdatingSetSelection = false;
            }

            RefreshFilterState();
        }

        private void ApplySort()
        {
            if (SelectedSortMode == "Custom")
                return;

            if (SelectedSortMode == "Revit Sheet Schedule")
            {
                ApplyRevitScheduleSort();
                return;
            }

            var items = AllSheets.ToList();
            items = SelectedSortMode == "Name"
                ? items.OrderBy(s => s.SheetName, new NaturalStringComparer()).ToList()
                : items.OrderBy(s => s.SheetNumber, new NaturalStringComparer()).ToList();

            ReplaceSheets(items);
        }

        private void ApplyRevitScheduleSort()
        {
            var items = AllSheets.ToList();
            if (items.Count == 0)
                return;

            var orderedNumbers = _sheetScheduleService.GetOrderedSheetNumbers(
                SelectedSheetSchedule?.ElementIdValue,
                items.Select(i => i.Model));

            if (orderedNumbers.Count == 0)
            {
                Trace.WriteLine("[SheetSchedule] No usable schedule order found; falling back to sheet number order");
                items = items.OrderBy(s => s.SheetNumber, new NaturalStringComparer()).ToList();
            }
            else
            {
                var rank = orderedNumbers
                    .Select((number, index) => new { number, index })
                    .GroupBy(x => x.number, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First().index, StringComparer.OrdinalIgnoreCase);

                items = items
                    .OrderBy(s => rank.TryGetValue(s.SheetNumber ?? "", out var order) ? order : int.MaxValue)
                    .ThenBy(s => s.SheetNumber, new NaturalStringComparer())
                    .ToList();
            }

            ReplaceSheets(items);
        }

        private void ReplaceSheets(IEnumerable<SheetItemViewModel> items)
        {
            AllSheets.Clear();
            foreach (var item in items)
                AllSheets.Add(item);
        }

        private void OnSheetSelectionChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(SheetItemViewModel.IsSelected))
                return;

            UpdateSelectedCount();
            RefreshDerivedState();
        }

        private void UpdateSelectedCount()
        {
            SelectedCount = AllSheets.Count(s => s.IsSelected);
        }

        private void SelectVisibleSheets(bool isSelected)
        {
            foreach (var sheet in GetVisibleSheets())
                sheet.IsSelected = isSelected;

            UpdateSelectedCount();
            RefreshDerivedState();
        }

        private void ReorderSelectedSheets()
        {
            var selected = AllSheets.Where(s => s.IsSelected).ToList();
            if (selected.Count < 2)
                return;

            var dialog = new ReorderSheetsDialog(selected)
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() != true)
                return;

            var selectedById = selected.ToDictionary(
                s => ViewSheetSetService.GetElementIdValue(s.ElementId),
                s => s,
                StringComparer.OrdinalIgnoreCase);

            var orderedSelected = dialog.OrderedIds
                .Where(id => selectedById.ContainsKey(id))
                .Select(id => selectedById[id])
                .ToList();

            if (orderedSelected.Count != selected.Count)
                return;

            var queue = new Queue<SheetItemViewModel>(orderedSelected);
            var reordered = AllSheets
                .Select(item => item.IsSelected ? queue.Dequeue() : item)
                .ToList();

            ReplaceSheets(reordered);
            SelectedSortMode = "Custom";
            RefreshFilterState();
        }

        private async Task SaveSelectedSheetsAsSetAsync()
        {
            var selected = AllSheets.Where(s => s.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Select at least one sheet before saving a Sheet Set.",
                    "No Sheets Selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var dialog = new SaveSheetSetDialog(selected.Count)
            {
                Owner = GetActiveWindow()
            };

            if (dialog.ShowDialog() != true)
                return;

            var setName = dialog.SetName;
            var selectedIds = selected.Select(s => s.ElementId).ToList();

            try
            {
                var exists = await _revitThreadService.RunOnRevitThreadAsync(app =>
                    _viewSheetSetService.SheetSetExists(setName));

                var replaceExisting = false;
                if (exists)
                {
                    var result = MessageBox.Show(
                        "A Sheet Set named '" + setName + "' already exists.\n\nReplace it?",
                        "Sheet Set Exists",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (result != MessageBoxResult.Yes)
                        return;

                    replaceExisting = true;
                }

                var sets = await _revitThreadService.RunOnRevitThreadAsync(app =>
                {
                    _viewSheetSetService.CreateSheetSet(setName, selectedIds, replaceExisting);
                    return _viewSheetSetService.GetSheetSets();
                });

                LoadViewSheetSets(sets);
                SelectOnlySheetSet(setName);
                StatusMessage = "Saved Sheet Set: " + setName;

                MessageBox.Show(
                    "Sheet Set '" + setName + "' saved successfully.\n\nContains " + selected.Count + " sheet(s).",
                    "Sheet Set Saved",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[ViewSheetSet] Save failed: " + ex);
                MessageBox.Show(
                    "Failed to save Sheet Set:\n" + ex.Message,
                    "Sheet Set Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SelectOnlySheetSet(string setName)
        {
            try
            {
                _isUpdatingSetSelection = true;
                foreach (var set in AvailableViewSheetSets)
                    set.IsSelected = !set.IsBuiltIn && string.Equals(set.Name, setName, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                _isUpdatingSetSelection = false;
            }

            RefreshFilterState();
        }

        private class InitialLoadData
        {
            public List<SheetInfo> Sheets { get; set; }
            public List<SheetScheduleInfo> Schedules { get; set; }
            public List<ViewSheetSetInfo> ViewSheetSets { get; set; }
            public List<string> ExportSetups { get; set; }
            public string ProjectName { get; set; }
            public string ProjectNumber { get; set; }
        }
    }
}
