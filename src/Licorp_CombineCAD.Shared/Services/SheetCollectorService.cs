using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Collects sheet metadata. The initial grid load stays intentionally light;
    /// export-only checks are hydrated only for selected sheets at export time.
    /// </summary>
    public class SheetCollectorService
    {
        private readonly Document _document;

        public SheetCollectorService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Get all sheets in the document as SheetInfo objects
        /// </summary>
        public List<SheetInfo> GetAllSheets()
        {
            var sheets = new List<SheetInfo>();
            var timer = Stopwatch.StartNew();

            try
            {
                var titleBlocksBySheetId = GetTitleBlocksBySheetId();
                var viewSheets = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewSheet))
                    .WhereElementIsNotElementType()
                    .Cast<ViewSheet>()
                    .OrderBy(s => s.SheetNumber)
                    .ToList();

                foreach (ViewSheet viewSheet in viewSheets)
                {
                    var sheetInfo = CreateBasicSheetInfo(viewSheet);
                    titleBlocksBySheetId.TryGetValue(GetElementIdValue(viewSheet.Id), out var titleBlock);
                    ApplyPaperSize(sheetInfo, titleBlock);
                    sheets.Add(sheetInfo);
                }

                Trace.WriteLine($"[CombineCAD] Collected {sheets.Count} sheet rows in {timer.ElapsedMilliseconds}ms");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Error collecting sheets: {ex.Message}");
            }

            return sheets;
        }

        /// <summary>
        /// Fill export-only details for selected sheets. This keeps form opening fast.
        /// </summary>
        public void HydrateSheetsForExport(IList<SheetInfo> sheets)
        {
            if (sheets == null || sheets.Count == 0)
                return;

            var timer = Stopwatch.StartNew();
            var requestedIds = new HashSet<string>(
                sheets.Select(s => GetElementIdValue(s.ElementId)).Where(id => !string.IsNullOrWhiteSpace(id)),
                StringComparer.OrdinalIgnoreCase);
            var titleBlocksBySheetId = GetTitleBlocksBySheetId(requestedIds);

            foreach (var sheet in sheets)
            {
                var viewSheet = _document.GetElement(sheet.ElementId) as ViewSheet;
                if (viewSheet == null)
                    continue;

                sheet.SheetNumber = viewSheet.SheetNumber ?? sheet.SheetNumber ?? "";
                sheet.SheetName = viewSheet.Name ?? sheet.SheetName ?? "";
                if (string.IsNullOrWhiteSpace(sheet.PaperSize))
                {
                    titleBlocksBySheetId.TryGetValue(GetElementIdValue(viewSheet.Id), out var titleBlock);
                    ApplyPaperSize(sheet, titleBlock);
                }

                ApplyViewportPresence(sheet, viewSheet);
            }

            Trace.WriteLine($"[CombineCAD] Hydrated {sheets.Count} selected sheets in {timer.ElapsedMilliseconds}ms");
        }

        private SheetInfo CreateBasicSheetInfo(ViewSheet viewSheet)
        {
            return new SheetInfo
            {
                ElementId = viewSheet.Id,
                SheetNumber = viewSheet.SheetNumber ?? "",
                SheetName = viewSheet.Name ?? "",
                PaperSize = "",
                IsSelected = false
            };
        }

        private Dictionary<string, FamilyInstance> GetTitleBlocksBySheetId(ISet<string> allowedSheetIds = null)
        {
            var result = new Dictionary<string, FamilyInstance>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var titleBlocks = new FilteredElementCollector(_document)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>();

                foreach (var titleBlock in titleBlocks)
                {
                    var sheetId = GetElementIdValue(titleBlock.OwnerViewId);
                    if (string.IsNullOrWhiteSpace(sheetId))
                        continue;
                    if (allowedSheetIds != null && !allowedSheetIds.Contains(sheetId))
                        continue;
                    if (!result.ContainsKey(sheetId))
                        result.Add(sheetId, titleBlock);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[CombineCAD] Title block map failed: " + ex.Message);
            }

            return result;
        }

        private static void ApplyPaperSize(SheetInfo info, Element titleBlock)
        {
            if (info == null || titleBlock == null)
                return;

            try
            {
                var sizeParam = titleBlock.LookupParameter("Sheet Size")
                    ?? titleBlock.LookupParameter("Paper Size");
                var paperSize = NormalizePaperSize(sizeParam?.AsString());

                if (string.IsNullOrWhiteSpace(paperSize))
                    paperSize = NormalizePaperSize(BuildTitleBlockDescriptor(titleBlock));

                if (string.IsNullOrWhiteSpace(paperSize))
                    paperSize = DetectPaperSize(titleBlock);

                info.PaperSize = paperSize;
            }
            catch
            {
                info.PaperSize = "";
            }
        }

        private void ApplyViewportPresence(SheetInfo info, ViewSheet viewSheet)
        {
            try
            {
                var viewportIds = viewSheet.GetAllViewports();
                info.HasNoView = viewportIds == null || viewportIds.Count == 0;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Viewport check error for {info.SheetNumber}: {ex.Message}");
                info.HasNoView = false;
            }
        }

        private static string GetElementIdValue(ElementId id)
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

        private static string DetectPaperSize(Element titleBlock)
        {
            try
            {
                var width = titleBlock.get_Parameter(BuiltInParameter.SHEET_WIDTH)?.AsDouble() ?? 0;
                var height = titleBlock.get_Parameter(BuiltInParameter.SHEET_HEIGHT)?.AsDouble() ?? 0;

                if (width <= 0 || height <= 0)
                    return "";

                return ClassifyPaperSize(width * 304.8, height * 304.8);
            }
            catch
            {
                return "";
            }
        }

        private static string BuildTitleBlockDescriptor(Element titleBlock)
        {
            if (titleBlock == null)
                return "";

            try
            {
                var familyInstance = titleBlock as FamilyInstance;
                var symbolName = familyInstance?.Symbol?.Name ?? "";
                var familyName = familyInstance?.Symbol?.FamilyName ?? "";

                return string.Join(" ", new[] { titleBlock.Name ?? "", symbolName, familyName }
                    .Where(value => !string.IsNullOrWhiteSpace(value)));
            }
            catch
            {
                return titleBlock.Name ?? "";
            }
        }

        private static string NormalizePaperSize(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return "";

            var normalized = rawValue.Trim().ToUpperInvariant();
            var isoMatch = Regex.Match(normalized, @"\bA([0-4])\b", RegexOptions.CultureInvariant);
            if (isoMatch.Success)
                return "A" + isoMatch.Groups[1].Value;

            var mmMatch = Regex.Match(normalized, @"(?<width>\d{2,4}(?:[.,]\d+)?)\s*[Xx]\s*(?<height>\d{2,4}(?:[.,]\d+)?)");
            if (mmMatch.Success &&
                TryParseMillimeter(mmMatch.Groups["width"].Value, out var widthMm) &&
                TryParseMillimeter(mmMatch.Groups["height"].Value, out var heightMm))
            {
                return ClassifyPaperSize(widthMm, heightMm);
            }

            return normalized;
        }

        private static bool TryParseMillimeter(string rawValue, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(rawValue))
                return false;

            var normalized = rawValue.Replace(',', '.');
            return double.TryParse(
                normalized,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out value);
        }

        private static string ClassifyPaperSize(double widthMm, double heightMm)
        {
            if (IsCloseToSize(widthMm, heightMm, 841, 1189))
                return "A0";
            if (IsCloseToSize(widthMm, heightMm, 594, 841))
                return "A1";
            if (IsCloseToSize(widthMm, heightMm, 420, 594))
                return "A2";
            if (IsCloseToSize(widthMm, heightMm, 297, 420))
                return "A3";
            if (IsCloseToSize(widthMm, heightMm, 210, 297))
                return "A4";

            return string.Format("{0:0} x {1:0} mm", Math.Max(widthMm, heightMm), Math.Min(widthMm, heightMm));
        }

        private static bool IsCloseToSize(double widthMm, double heightMm, double standardWidthMm, double standardHeightMm)
        {
            const double toleranceMm = 10.0;
            return (Math.Abs(widthMm - standardWidthMm) <= toleranceMm &&
                    Math.Abs(heightMm - standardHeightMm) <= toleranceMm) ||
                   (Math.Abs(widthMm - standardHeightMm) <= toleranceMm &&
                    Math.Abs(heightMm - standardWidthMm) <= toleranceMm);
        }
    }
}
