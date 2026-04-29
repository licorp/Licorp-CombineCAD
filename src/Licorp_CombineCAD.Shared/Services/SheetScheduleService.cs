using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class SheetScheduleService
    {
        private readonly Document _document;

        public SheetScheduleService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public List<SheetScheduleInfo> GetSheetListSchedules()
        {
            var result = new List<SheetScheduleInfo>();

            try
            {
                var schedules = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .Where(IsSheetListSchedule)
                    .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase);

                foreach (var schedule in schedules)
                {
                    result.Add(new SheetScheduleInfo
                    {
                        ElementIdValue = GetElementIdValue(schedule.Id).ToString(),
                        Name = schedule.Name
                    });
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[SheetSchedule] Failed to collect schedules: " + ex.Message);
            }

            return result;
        }

        public List<string> GetOrderedSheetNumbers(string scheduleIdValue, IEnumerable<SheetInfo> candidateSheets)
        {
            var selectedSheets = (candidateSheets ?? Enumerable.Empty<SheetInfo>())
                .Where(s => s != null)
                .ToList();

            if (selectedSheets.Count == 0 || string.IsNullOrWhiteSpace(scheduleIdValue))
                return new List<string>();

            var sheetNumbers = new HashSet<string>(
                selectedSheets.Select(s => s.SheetNumber ?? ""),
                StringComparer.OrdinalIgnoreCase);

            try
            {
                var schedule = GetSchedule(scheduleIdValue);
                if (schedule == null)
                    return new List<string>();

                var tableData = schedule.GetTableData();
                var body = tableData?.GetSectionData(SectionType.Body);
                if (body == null)
                    return new List<string>();

                var ordered = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int row = 0; row < body.NumberOfRows; row++)
                {
                    for (int col = 0; col < body.NumberOfColumns; col++)
                    {
                        var cellText = "";
                        try
                        {
                            cellText = schedule.GetCellText(SectionType.Body, row, col);
                        }
                        catch
                        {
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(cellText))
                            continue;

                        var match = sheetNumbers.FirstOrDefault(n =>
                            !string.IsNullOrWhiteSpace(n) &&
                            string.Equals(n.Trim(), cellText.Trim(), StringComparison.OrdinalIgnoreCase));

                        if (!string.IsNullOrWhiteSpace(match) && seen.Add(match))
                        {
                            ordered.Add(match);
                            break;
                        }
                    }
                }

                Trace.WriteLine($"[SheetSchedule] Schedule '{schedule.Name}' returned {ordered.Count} ordered sheet(s)");
                return ordered;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[SheetSchedule] Failed to read schedule order: " + ex.Message);
                return new List<string>();
            }
        }

        private ViewSchedule GetSchedule(string scheduleIdValue)
        {
            if (!long.TryParse(scheduleIdValue, out var idValue))
                return null;

            var elementId = CreateElementId(idValue);
            return elementId == null ? null : _document.GetElement(elementId) as ViewSchedule;
        }

        private bool IsSheetListSchedule(ViewSchedule schedule)
        {
            if (schedule == null || schedule.IsTemplate)
                return false;

            try
            {
                var definition = schedule.Definition;
                if (definition == null)
                    return false;

                return GetElementIdValue(definition.CategoryId) == (long)BuiltInCategory.OST_Sheets;
            }
            catch
            {
                return false;
            }
        }

        private static long GetElementIdValue(ElementId id)
        {
            if (id == null)
                return 0;

            try
            {
                var valueProperty = typeof(ElementId).GetProperty("Value");
                if (valueProperty != null)
                    return Convert.ToInt64(valueProperty.GetValue(id, null));

                var integerValueProperty = typeof(ElementId).GetProperty("IntegerValue");
                if (integerValueProperty != null)
                    return Convert.ToInt64(integerValueProperty.GetValue(id, null));
            }
            catch
            {
            }

            return 0;
        }

        private static ElementId CreateElementId(long value)
        {
            try
            {
                var longCtor = typeof(ElementId).GetConstructor(new[] { typeof(long) });
                if (longCtor != null)
                    return (ElementId)longCtor.Invoke(new object[] { value });

                var intCtor = typeof(ElementId).GetConstructor(new[] { typeof(int) });
                if (intCtor != null)
                    return (ElementId)intCtor.Invoke(new object[] { Convert.ToInt32(value) });
            }
            catch
            {
            }

            return null;
        }
    }
}
