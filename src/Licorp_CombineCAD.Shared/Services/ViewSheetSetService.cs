using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class ViewSheetSetService
    {
        private readonly Document _document;

        public ViewSheetSetService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public List<ViewSheetSetInfo> GetSheetSets()
        {
            var result = new List<ViewSheetSetInfo>();

            try
            {
                var allSheetsSet = new ViewSheetSetInfo
                {
                    Name = "All Sheets",
                    IsBuiltIn = true
                };
                result.Add(allSheetsSet);

                var savedSets = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewSheetSet))
                    .Cast<ViewSheetSet>()
                    .Where(s => s != null && !string.IsNullOrWhiteSpace(s.Name))
                    .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase);

                foreach (var savedSet in savedSets)
                {
                    var info = new ViewSheetSetInfo
                    {
                        Name = savedSet.Name,
                        IsBuiltIn = false
                    };

                    try
                    {
                        if (savedSet.Views != null && !savedSet.Views.IsEmpty)
                        {
                            foreach (View view in savedSet.Views)
                            {
                                if (view is ViewSheet sheet && !sheet.IsTemplate)
                                {
                                    var idValue = GetElementIdValue(sheet.Id);
                                    if (!string.IsNullOrWhiteSpace(idValue) &&
                                        !info.SheetIdValues.Contains(idValue, StringComparer.OrdinalIgnoreCase))
                                    {
                                        info.SheetIdValues.Add(idValue);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine("[ViewSheetSet] Failed reading set '" + savedSet.Name + "': " + ex.Message);
                    }

                    result.Add(info);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[ViewSheetSet] Failed collecting sets: " + ex.Message);
            }

            return result;
        }

        public bool SheetSetExists(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            return new FilteredElementCollector(_document)
                .OfClass(typeof(ViewSheetSet))
                .Cast<ViewSheetSet>()
                .Any(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        public ViewSheetSet CreateSheetSet(string name, IEnumerable<ElementId> sheetIds, bool replaceExisting)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Sheet Set name cannot be empty.", nameof(name));

            var ids = (sheetIds ?? Enumerable.Empty<ElementId>())
                .Where(id => id != null)
                .ToList();

            if (ids.Count == 0)
                throw new ArgumentException("Select at least one sheet.", nameof(sheetIds));

            using (var transaction = new Transaction(_document, "Create Sheet Set"))
            {
                transaction.Start();

                try
                {
                    var existing = new FilteredElementCollector(_document)
                        .OfClass(typeof(ViewSheetSet))
                        .Cast<ViewSheetSet>()
                        .FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

                    if (existing != null)
                    {
                        if (!replaceExisting)
                            throw new InvalidOperationException("Sheet Set '" + name + "' already exists.");

                        _document.Delete(existing.Id);
                    }

                    var viewSet = new ViewSet();
                    foreach (var id in ids)
                    {
                        var view = _document.GetElement(id) as View;
                        if (view != null && view.CanBePrinted && !view.IsTemplate)
                            viewSet.Insert(view);
                    }

                    if (viewSet.IsEmpty)
                        throw new InvalidOperationException("No printable sheets were selected.");

                    var printManager = _document.PrintManager;
                    printManager.PrintRange = PrintRange.Select;
                    var viewSheetSetting = printManager.ViewSheetSetting;
                    viewSheetSetting.CurrentViewSheetSet.Views = viewSet;
                    viewSheetSetting.SaveAs(name);

                    transaction.Commit();

                    return new FilteredElementCollector(_document)
                        .OfClass(typeof(ViewSheetSet))
                        .Cast<ViewSheetSet>()
                        .FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                }
                catch
                {
                    transaction.RollBack();
                    throw;
                }
            }
        }

        public static string GetElementIdValue(ElementId id)
        {
            if (id == null)
                return "";

            try
            {
                var valueProperty = typeof(ElementId).GetProperty("Value");
                if (valueProperty != null)
                    return Convert.ToInt64(valueProperty.GetValue(id, null)).ToString();

                var integerValueProperty = typeof(ElementId).GetProperty("IntegerValue");
                if (integerValueProperty != null)
                    return Convert.ToInt64(integerValueProperty.GetValue(id, null)).ToString();
            }
            catch
            {
            }

            return "";
        }
    }
}
