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
                            info.Area = (outline.MaximumPoint.X - outline.MinimumPoint.X) *
                                        (outline.MaximumPoint.Y - outline.MinimumPoint.Y);
                        }
                    }
                    catch { }

                    viewports.Add(info);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SmartScale] Error getting viewports: {ex.Message}");
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
                Debug.WriteLine($"[SmartScale] Error finding title block: {ex.Message}");
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
                    Debug.WriteLine($"[SmartScale] Title block not found for sheet {sheet.SheetNumber}");
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
                    Debug.WriteLine($"[SmartScale] Set {scaleText} on title block for sheet {sheet.SheetNumber}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SmartScale] Error applying scale: {ex.Message}");
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
                    Debug.WriteLine($"[SmartScale] Restored original scale for sheet {sheet.SheetNumber}");
                }

                _originalValues.Remove(sheet.Id);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SmartScale] Error restoring scale: {ex.Message}");
            }
        }
    }

    public class ViewportInfo
    {
        public ElementId ElementId { get; set; }
        public ElementId ViewId { get; set; }
        public string ViewName { get; set; }
        public int Scale { get; set; }
        public string ScaleText { get; set; }
        public double Area { get; set; }
    }
}