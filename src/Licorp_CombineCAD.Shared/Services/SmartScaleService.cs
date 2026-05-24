using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class SmartScaleService
    {
        private readonly Document _document;
        private readonly Dictionary<ElementId, string> _originalValues = new Dictionary<ElementId, string>();

        public SmartScaleService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public List<ViewportInfo> GetViewportsOnSheet(ViewSheet sheet)
        {
            var viewports = new List<ViewportInfo>();

            try
            {
                var viewportIds = sheet.GetAllViewports();
                if (viewportIds == null || viewportIds.Count == 0)
                    return viewports;

                foreach (ElementId vpId in viewportIds)
                {
                    var viewport = _document.GetElement(vpId) as Viewport;
                    if (viewport == null) continue;

                    var view = _document.GetElement(viewport.ViewId) as View;
                    if (view == null) continue;

                    var info = new ViewportInfo
                    {
                        ElementId = vpId,
                        ViewId = viewport.ViewId,
                        ViewName = view.Name,
                        Scale = view.Scale,
                        ScaleText = $"1:{view.Scale}"
                    };

                    try
                    {
                        var outline = viewport.GetBoxOutline();
                        if (outline != null)
                        {
                            info.Width = outline.MaximumPoint.X - outline.MinimumPoint.X;
                            info.Height = outline.MaximumPoint.Y - outline.MinimumPoint.Y;
                            info.Area = info.Width * info.Height;
                        }
                    }
                    catch { }

                    viewports.Add(info);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error getting viewports: {ex.Message}");
            }

            return viewports;
        }

        public string FormatScale(int scale)
        {
            if (scale <= 0) return "As Indicated";
            return $"1:{scale}";
        }

        private ElementId FindTitleBlock(ViewSheet sheet)
        {
            try
            {
                var titleBlock = new FilteredElementCollector(_document, sheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .OfClass(typeof(FamilyInstance))
                    .FirstOrDefault();
                return titleBlock?.Id;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error finding title block: {ex.Message}");
            }

            return null;
        }

        public bool ApplySmartScale(ViewSheet sheet, Transaction trans)
        {
            try
            {
                var viewports = GetViewportsOnSheet(sheet);
                if (viewports.Count == 0) return false;

                var primaryScale = viewports.OrderByDescending(v => v.Area).First().Scale;
                var scaleText = FormatScale(primaryScale);

                var titleBlockId = FindTitleBlock(sheet);
                if (titleBlockId == null)
                {
                    Trace.WriteLine($"[SmartScale] Title block not found for sheet {sheet.SheetNumber}");
                    return false;
                }

                var titleBlock = _document.GetElement(titleBlockId);
                if (titleBlock == null) return false;

                var scaleParam = titleBlock.LookupParameter("Scale") ??
                                 titleBlock.LookupParameter("Drawing Scale") ??
                                 titleBlock.get_Parameter(BuiltInParameter.SHEET_SCALE);

                if (scaleParam != null && !scaleParam.IsReadOnly)
                {
                    _originalValues[sheet.Id] = scaleParam.AsString();
                    scaleParam.Set(scaleText);
                    Trace.WriteLine($"[SmartScale] Set {scaleText} on title block for sheet {sheet.SheetNumber}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error applying scale: {ex.Message}");
                return false;
            }
        }

        public void RestoreOriginalScale(ViewSheet sheet, Transaction trans)
        {
            try
            {
                if (!_originalValues.TryGetValue(sheet.Id, out var originalValue))
                    return;

                var titleBlockId = FindTitleBlock(sheet);
                if (titleBlockId == null) return;

                var titleBlock = _document.GetElement(titleBlockId);
                if (titleBlock == null) return;

                var scaleParam = titleBlock.LookupParameter("Scale") ??
                                 titleBlock.LookupParameter("Drawing Scale") ??
                                 titleBlock.get_Parameter(BuiltInParameter.SHEET_SCALE);

                if (scaleParam != null && !scaleParam.IsReadOnly)
                {
                    scaleParam.Set(originalValue);
                    Trace.WriteLine($"[SmartScale] Restored original scale for sheet {sheet.SheetNumber}");
                }

                _originalValues.Remove(sheet.Id);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error restoring scale: {ex.Message}");
            }
        }

        // ===== #1: As Indicated Fix (Mixed Scales) =====

        /// <summary>
        /// Kiem tra sheet co nhieu viewport voi scale khac nhau khong
        /// </summary>
        public bool HasMultipleScales(ViewSheet sheet)
        {
            try
            {
                var viewports = GetViewportsOnSheet(sheet);
                if (viewports.Count <= 1) return false;

                var distinctScales = viewports
                    .Select(v => v.Scale)
                    .Distinct()
                    .Count();

                return distinctScales > 1;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] HasMultipleScales error: {ex.Message}");
                return false;
            }
        }

        // ===== #2: Smart Scale Nang Cao =====

        /// <summary>
        /// Phan tich chi tat ca viewport tren sheet
        /// </summary>
        public ScaleAnalysis AnalyzeScales(ViewSheet sheet)
        {
            var analysis = new ScaleAnalysis();

            try
            {
                var viewports = GetViewportsOnSheet(sheet);
                if (viewports.Count == 0)
                {
                    analysis.RecommendedScale = "N/A";
                    return analysis;
                }

                double totalArea = viewports.Sum(v => v.Area);

                foreach (var vp in viewports)
                {
                    var info = new ViewportScaleInfo
                    {
                        ViewName = vp.ViewName,
                        Scale = vp.Scale,
                        ScaleText = FormatScale(vp.Scale),
                        Area = vp.Area,
                        AreaPercentage = totalArea > 0
                            ? (vp.Area / totalArea * 100)
                            : 0,
                        IsPrimary = vp.Area == viewports.Max(v => v.Area)
                    };

                    analysis.ViewportScales.Add(info);
                }

                analysis.PrimaryViewport = analysis.ViewportScales
                    .First(v => v.IsPrimary);

                analysis.DistinctScaleCount = analysis.ViewportScales
                    .Select(v => v.Scale)
                    .Distinct()
                    .Count();

                analysis.HasMultipleScales = analysis.DistinctScaleCount > 1;

                analysis.RecommendedScale = analysis.HasMultipleScales
                    ? "As Indicated"
                    : $"1:{analysis.PrimaryViewport.Scale}";

                Trace.WriteLine($"[SmartScale] Sheet {sheet.SheetNumber}: " +
                    $"{analysis.ViewportScales.Count} viewports, " +
                    $"{analysis.DistinctScaleCount} distinct scales, " +
                    $"recommended: {analysis.RecommendedScale}");

                foreach (var vp in analysis.ViewportScales)
                {
                    Trace.WriteLine($"  - {vp.ViewName}: {vp.ScaleText} " +
                        $"({vp.AreaPercentage:F1}% area)" +
                        (vp.IsPrimary ? " [PRIMARY]" : ""));
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] AnalyzeScales error: {ex.Message}");
            }

            return analysis;
        }
    }

    public class ViewportInfo
    {
        public ElementId ElementId { get; set; }
        public ElementId ViewId { get; set; }
        public string ViewName { get; set; }
        public int Scale { get; set; }
        public string ScaleText { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Area { get; set; }
    }

    /// <summary>
    /// Thong tin phan tich scale cua sheet
    /// </summary>
    public class ScaleAnalysis
    {
        public List<ViewportScaleInfo> ViewportScales { get; set; } = new List<ViewportScaleInfo>();
        public bool HasMultipleScales { get; set; }
        public string RecommendedScale { get; set; }
        public ViewportScaleInfo PrimaryViewport { get; set; }
        public int DistinctScaleCount { get; set; }
    }

    /// <summary>
    /// Thong tin scale cua tung viewport
    /// </summary>
    public class ViewportScaleInfo
    {
        public string ViewName { get; set; }
        public int Scale { get; set; }
        public string ScaleText { get; set; }
        public double Area { get; set; }
        public double AreaPercentage { get; set; }
        public bool IsPrimary { get; set; }
    }
}