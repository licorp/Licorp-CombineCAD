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

        public int GetPrimaryViewScale(ViewSheet sheet)
        {
            var viewports = GetViewportsOnSheet(sheet);
            if (viewports.Count == 0) return 0;

            var primary = viewports.OrderByDescending(v => v.Area).First();
            return primary.Scale;
        }

        public string FormatScale(int scale)
        {
            if (scale <= 0) return "As Indicated";
            return $"1:{scale}";
        }

        // ===== Phase 2: Sheet-Ratio-Based Scaling =====

        /// <summary>
        /// Get paper size info from Revit sheet outline.
        /// Returns width and height in millimeters.
        /// </summary>
        public PaperSizeInfo GetPaperSizeInfo(ViewSheet sheet)
        {
            try
            {
                var outline = sheet.Outline;
                double widthFeet = outline.Max.U - outline.Min.U;
                double heightFeet = outline.Max.V - outline.Min.V;

                // Convert feet to mm (1 foot = 304.8 mm)
                double widthMm = widthFeet * 304.8;
                double heightMm = heightFeet * 304.8;

                return new PaperSizeInfo
                {
                    WidthMm = widthMm,
                    HeightMm = heightMm,
                    SizeName = ClassifyPaperSize(widthMm, heightMm)
                };
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error getting paper size: {ex.Message}");
                return new PaperSizeInfo { WidthMm = 0, HeightMm = 0, SizeName = "Unknown" };
            }
        }

        /// <summary>
        /// Calculate the ratio between sheet size and primary viewport size.
        /// Used for Sheet-Ratio-Based Scaling (MLabs Option 2).
        /// 
        /// Formula: ratio = min(sheetWidth/vpWidth, sheetHeight/vpHeight)
        /// 
        /// When ratio > 1: viewport is smaller than sheet, content needs to be scaled up
        /// When ratio = 1: viewport fits exactly
        /// When ratio < 1: viewport is larger than sheet, content needs to be scaled down
        /// </summary>
        public double CalculateSheetRatio(ViewSheet sheet)
        {
            try
            {
                var paperInfo = GetPaperSizeInfo(sheet);
                if (paperInfo.WidthMm <= 0 || paperInfo.HeightMm <= 0)
                {
                    Trace.WriteLine($"[SmartScale] Invalid paper size for sheet {sheet.SheetNumber}");
                    return 1.0;
                }

                var viewports = GetViewportsOnSheet(sheet);
                if (viewports.Count == 0)
                {
                    Trace.WriteLine($"[SmartScale] No viewports on sheet {sheet.SheetNumber}, using ratio 1.0");
                    return 1.0;
                }

                // Find the primary (largest) viewport
                var primary = viewports.OrderByDescending(v => v.Area).First();

                // Viewport dimensions are in Revit internal units (feet)
                double vpWidthMm = primary.Width * 304.8;
                double vpHeightMm = primary.Height * 304.8;

                if (vpWidthMm <= 0 || vpHeightMm <= 0)
                {
                    Trace.WriteLine($"[SmartScale] Invalid viewport size for sheet {sheet.SheetNumber}");
                    return 1.0;
                }

                // Calculate ratio for both dimensions
                double ratioW = paperInfo.WidthMm / vpWidthMm;
                double ratioH = paperInfo.HeightMm / vpHeightMm;

                // Use the smaller ratio to ensure content fits in both dimensions
                double ratio = Math.Min(ratioW, ratioH);

                Trace.WriteLine($"[SmartScale] Sheet {sheet.SheetNumber}: " +
                    $"Paper={paperInfo.WidthMm:F0}x{paperInfo.HeightMm:F0}mm, " +
                    $"Viewport={vpWidthMm:F0}x{vpHeightMm:F0}mm, " +
                    $"Ratio={ratio:F4} (W={ratioW:F4}, H={ratioH:F4})");

                return ratio;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error calculating sheet ratio: {ex.Message}");
                return 1.0;
            }
        }

        /// <summary>
        /// Apply sheet-ratio scale to title block parameter.
        /// The ratio is formatted as a scale string (e.g., "1:1.5").
        /// </summary>
        public bool ApplySheetRatioScale(ViewSheet sheet, Transaction trans)
        {
            try
            {
                double ratio = CalculateSheetRatio(sheet);
                if (Math.Abs(ratio - 1.0) < 0.001)
                {
                    Trace.WriteLine($"[SmartScale] Sheet {sheet.SheetNumber} ratio is ~1.0, no scaling needed");
                    return false;
                }

                // Format ratio as scale text
                string scaleText;
                if (ratio >= 1.0)
                {
                    // Scale up: 1:X where X < 1 (e.g., 1:0.707)
                    double scaleValue = 1.0 / ratio;
                    scaleText = $"1:{scaleValue:F3}";
                }
                else
                {
                    // Scale down: X:1 where X > 1 (e.g., 1.5:1)
                    double scaleValue = ratio;
                    scaleText = $"{scaleValue:F3}:1";
                }

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
                    Trace.WriteLine($"[SmartScale] Applied sheet ratio {scaleText} for sheet {sheet.SheetNumber}");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] Error applying sheet ratio: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Classify paper size from mm dimensions to standard name (A0-A4)
        /// </summary>
        private static string ClassifyPaperSize(double widthMm, double heightMm)
        {
            double maxDim = Math.Max(widthMm, heightMm);
            double minDim = Math.Min(widthMm, heightMm);

            if (IsCloseTo(maxDim, 1189) && IsCloseTo(minDim, 841)) return "A0";
            if (IsCloseTo(maxDim, 841) && IsCloseTo(minDim, 594)) return "A1";
            if (IsCloseTo(maxDim, 594) && IsCloseTo(minDim, 420)) return "A2";
            if (IsCloseTo(maxDim, 420) && IsCloseTo(minDim, 297)) return "A3";
            if (IsCloseTo(maxDim, 297) && IsCloseTo(minDim, 210)) return "A4";

            return $"{maxDim:F0}x{minDim:F0}mm";
        }

        private static bool IsCloseTo(double value, double target, double tolerance = 10.0)
        {
            return Math.Abs(value - target) <= tolerance;
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

        /// <summary>
        /// Lay text hien thi cho scale tren title block
        /// - Neu 1 scale duy nhat: "1:50"
        /// - Neu nhieu scale: "As Indicated"
        /// </summary>
        public string GetScaleDisplayText(ViewSheet sheet)
        {
            if (HasMultipleScales(sheet))
            {
                Trace.WriteLine($"[SmartScale] Sheet {sheet.SheetNumber} has multiple scales -> 'As Indicated'");
                return "As Indicated";
            }

            var primaryScale = GetPrimaryViewScale(sheet);
            return FormatScale(primaryScale);
        }

        /// <summary>
        /// Ap dung scale dung cho title block
        /// - Neu 1 scale: ghi "1:50"
        /// - Neu nhieu scale: ghi "As Indicated"
        /// </summary>
        public bool ApplyCorrectScale(ViewSheet sheet, Transaction trans)
        {
            try
            {
                string scaleText = GetScaleDisplayText(sheet);

                var titleBlockId = FindTitleBlock(sheet);
                if (titleBlockId == null)
                {
                    Trace.WriteLine($"[SmartScale] Title block not found for {sheet.SheetNumber}");
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
                    Trace.WriteLine($"[SmartScale] Applied '{scaleText}' for sheet {sheet.SheetNumber}");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SmartScale] ApplyCorrectScale error: {ex.Message}");
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

    /// <summary>
    /// Paper size information from Revit sheet outline
    /// </summary>
    public class PaperSizeInfo
    {
        public double WidthMm { get; set; }
        public double HeightMm { get; set; }
        public string SizeName { get; set; }
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