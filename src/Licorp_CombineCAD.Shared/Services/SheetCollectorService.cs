using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Collects sheet metadata. The initial grid load stays intentionally light;
    /// heavier viewport/scale analysis is hydrated only for selected sheets at export time.
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
                ApplyRevision(sheet, viewSheet);

                if (string.IsNullOrWhiteSpace(sheet.PaperSize))
                {
                    titleBlocksBySheetId.TryGetValue(GetElementIdValue(viewSheet.Id), out var titleBlock);
                    ApplyPaperSize(sheet, titleBlock);
                }

                AnalyzeViewports(sheet, viewSheet);
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
                Revision = "",
                PaperSize = "",
                ScaleText = "",
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

        private static void ApplyRevision(SheetInfo info, ViewSheet viewSheet)
        {
            try
            {
                var revParam = viewSheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
                info.Revision = revParam?.AsString() ?? "";
            }
            catch
            {
                info.Revision = "";
            }
        }

        private static void ApplyPaperSize(SheetInfo info, Element titleBlock)
        {
            if (info == null || titleBlock == null)
                return;

            try
            {
                var sizeParam = titleBlock.LookupParameter("Sheet Size")
                    ?? titleBlock.LookupParameter("Paper Size");
                info.PaperSize = sizeParam?.AsString() ?? "";
                if (string.IsNullOrWhiteSpace(info.PaperSize))
                    info.PaperSize = DetectPaperSize(titleBlock);
            }
            catch
            {
                info.PaperSize = "";
            }
        }

        private void AnalyzeViewports(SheetInfo info, ViewSheet viewSheet)
        {
            try
            {
                var viewportIds = viewSheet.GetAllViewports();
                info.HasNoView = viewportIds == null || viewportIds.Count == 0;

                if (info.HasNoView)
                {
                    info.ViewScales = new List<int>();
                    info.PrimaryScale = 0;
                    info.ScaleText = "No views";
                    return;
                }

                var scales = new List<int>();
                double maxArea = 0;
                int primaryScale = 0;

                foreach (ElementId vpId in viewportIds)
                {
                    var viewport = _document.GetElement(vpId) as Viewport;
                    if (viewport == null) continue;

                    var view = _document.GetElement(viewport.ViewId) as View;
                    if (view == null) continue;

                    int viewScale = view.Scale;
                    scales.Add(viewScale);

                    try
                    {
                        var outline = viewport.GetBoxOutline();
                        if (outline != null)
                        {
                            double area = (outline.MaximumPoint.X - outline.MinimumPoint.X)
                                        * (outline.MaximumPoint.Y - outline.MinimumPoint.Y);
                            if (area > maxArea)
                            {
                                maxArea = area;
                                primaryScale = viewScale;
                            }
                        }
                    }
                    catch
                    {
                        if (primaryScale == 0) primaryScale = viewScale;
                    }
                }

                info.ViewScales = scales;
                info.PrimaryScale = primaryScale;

                var uniqueScales = scales.Distinct().ToList();
                if (uniqueScales.Count == 1)
                    info.ScaleText = $"1:{uniqueScales[0]}";
                else if (uniqueScales.Count > 1)
                    info.ScaleText = "As Indicated";
                else
                    info.ScaleText = "-";
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Viewport analysis error for {info.SheetNumber}: {ex.Message}");
                info.ScaleText = "-";
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
