using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Collects sheets from the Revit document with viewport/scale information.
    /// Refactored from Export+ SheetBatchLoader.
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

            try
            {
                var collector = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewSheet))
                    .WhereElementIsNotElementType()
                    .Cast<ViewSheet>()
                    .OrderBy(s => s.SheetNumber);

                foreach (ViewSheet viewSheet in collector)
                {
                    var sheetInfo = CreateSheetInfo(viewSheet);
                    sheets.Add(sheetInfo);
                }

                Trace.WriteLine($"[CombineCAD] Collected {sheets.Count} sheets");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Error collecting sheets: {ex.Message}");
            }

            return sheets;
        }

        /// <summary>
        /// Create SheetInfo from a ViewSheet with viewport analysis
        /// </summary>
        private SheetInfo CreateSheetInfo(ViewSheet viewSheet)
        {
            var info = new SheetInfo
            {
                ElementId = viewSheet.Id,
                SheetNumber = viewSheet.SheetNumber ?? "",
                SheetName = viewSheet.Name ?? "",
                IsSelected = false,
            };

        // Get revision
        try
        {
            var revParam = viewSheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION);
            info.Revision = revParam?.AsString() ?? "";
        }
        catch { info.Revision = ""; }

        // Get paper size from title block
        try
        {
            var titleBlock = new FilteredElementCollector(_document, viewSheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilyInstance))
                .FirstOrDefault();
            if (titleBlock != null)
            {
                var sizeParam = titleBlock.LookupParameter("Sheet Size")
                    ?? titleBlock.LookupParameter("Paper Size");
                info.PaperSize = sizeParam?.AsString() ?? "";
            }
        }
        catch { info.PaperSize = ""; }

            // Get viewports and scales
            try
            {
                var viewportIds = viewSheet.GetAllViewports();
                info.HasNoView = viewportIds == null || viewportIds.Count == 0;

                if (!info.HasNoView)
                {
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

                        // Find primary view (largest viewport area)
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

                    // Build scale text
                    var uniqueScales = scales.Distinct().ToList();
                    if (uniqueScales.Count == 1)
                    {
                        info.ScaleText = $"1:{uniqueScales[0]}";
                    }
                    else if (uniqueScales.Count > 1)
                    {
                        info.ScaleText = "As Indicated";
                    }
                    else
                    {
                        info.ScaleText = "-";
                    }
                }
                else
                {
                    info.ScaleText = "No views";
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Viewport analysis error for {info.SheetNumber}: {ex.Message}");
                info.ScaleText = "-";
            }

            return info;
        }
    }
}
