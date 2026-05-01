using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class SheetPreflightService
    {
        private readonly Document _document;

        public SheetPreflightService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public SheetPreflightResult Analyze(IList<SheetInfo> sheets, ExportSettings settings)
        {
            var result = new SheetPreflightResult();

            if (sheets == null || sheets.Count == 0)
            {
                result.AddIssue(PreflightSeverity.Error, "", "", "Chưa chọn sheet nào.");
                return result;
            }

            CheckDuplicateOutputNames(sheets, settings, result);

            foreach (var sheet in sheets)
                AnalyzeSheet(sheet, result);

            LogResult(result);
            return result;
        }

        private void CheckDuplicateOutputNames(IList<SheetInfo> sheets, ExportSettings settings, SheetPreflightResult result)
        {
            var template = settings != null ? settings.FileNameTemplate : null;
            var groups = sheets
                .Select(s => new
                {
                    Sheet = s,
                    FileName = DwgExportService.GenerateFileName(s, template, _document)
                })
                .GroupBy(x => x.FileName ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var group in groups)
            {
                if (string.IsNullOrWhiteSpace(group.Key))
                {
                    foreach (var item in group)
                    {
                        result.AddIssue(
                            PreflightSeverity.Error,
                            item.Sheet.SheetNumber,
                            item.Sheet.SheetName,
                            "Tên file DWG tạo ra bị trống. Kiểm tra mẫu tên file.");
                    }
                    continue;
                }

                if (group.Count() <= 1)
                    continue;

                var sheetNames = string.Join(", ", group.Select(x => x.Sheet.SheetNumber));
                result.AddIssue(
                    PreflightSeverity.Error,
                    "",
                    "",
                    string.Format("Trùng lặp tên file DWG '{0}.dwg' từ các sheet: {1}", group.Key, sheetNames));
            }
        }

        private void AnalyzeSheet(SheetInfo sheet, SheetPreflightResult result)
        {
            if (sheet == null)
                return;

            var viewSheet = _document.GetElement(sheet.ElementId) as ViewSheet;
            if (viewSheet == null)
            {
                result.AddIssue(
                    PreflightSeverity.Error,
                    sheet.SheetNumber,
                    sheet.SheetName,
                    "Không tìm thấy ViewSheet này trong tài liệu Revit hiện tại.");
                return;
            }

            var titleBlockCount = CountTitleBlocks(viewSheet);
            var viewportCount = CountViewports(viewSheet);
            var rasterCount = CountElementsByCategory(viewSheet, BuiltInCategory.OST_RasterImages);
            var scheduleCount = CountSchedules(viewSheet);

            if (titleBlockCount == 0)
            {
                result.AddIssue(
                    PreflightSeverity.Warning,
                    sheet.SheetNumber,
                    sheet.SheetName,
                    "Không tìm thấy khung tên. Bản vẽ ghép có thể không có khung sheet.");
            }

            if (viewportCount == 0)
            {
                var hasOtherExportableContent = titleBlockCount > 0 || scheduleCount > 0 || rasterCount > 0;
                var severity = hasOtherExportableContent ? PreflightSeverity.Info : PreflightSeverity.Warning;
                result.AddIssue(
                    severity,
                    sheet.SheetNumber,
                    sheet.SheetName,
                    "Không phát hiện viewport mô hình. Sheet vẫn sẽ được xuất cho khung tên, bảng biểu, hình ảnh và chú thích.");
            }

            if (rasterCount > 0)
            {
                result.AddIssue(
                    PreflightSeverity.Warning,
                    sheet.SheetNumber,
                    sheet.SheetName,
                    string.Format("Phát hiện {0} hình ảnh raster. Revit có thể xuất chúng dưới dạng tham chiếu hình ảnh bên ngoài.", rasterCount));
            }
        }

        private int CountTitleBlocks(ViewSheet viewSheet)
        {
            try
            {
                return new FilteredElementCollector(_document, viewSheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .Count();
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Preflight] Title block count failed: " + ex.Message);
                return 0;
            }
        }

        private int CountViewports(ViewSheet viewSheet)
        {
            try
            {
                var viewportIds = viewSheet.GetAllViewports();
                return viewportIds == null ? 0 : viewportIds.Count;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Preflight] Viewport count failed: " + ex.Message);
                return 0;
            }
        }

        private int CountSchedules(ViewSheet viewSheet)
        {
            try
            {
                return new FilteredElementCollector(_document, viewSheet.Id)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .WhereElementIsNotElementType()
                    .Count();
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Preflight] Schedule count failed: " + ex.Message);
                return 0;
            }
        }

        private int CountElementsByCategory(ViewSheet viewSheet, BuiltInCategory category)
        {
            try
            {
                return new FilteredElementCollector(_document, viewSheet.Id)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .Count();
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[Preflight] Category count failed: " + category + ": " + ex.Message);
                return 0;
            }
        }

        private void LogResult(SheetPreflightResult result)
        {
            Trace.WriteLine("[Preflight] " + result.Summary);
            foreach (var issue in result.Issues)
            {
                Trace.WriteLine(string.Format(
                    "[Preflight] {0}: {1}: {2}",
                    issue.Severity,
                    issue.DisplayName,
                    issue.Message));
            }
        }
    }
}
