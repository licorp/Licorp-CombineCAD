using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class DwgExportService
    {
        private readonly Document _document;

        public DwgExportService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public List<string> GetAvailableExportSetups()
        {
            var setups = new List<string>();
            try
            {
                var collector = new FilteredElementCollector(_document)
                    .OfClass(typeof(ExportDWGSettings));

                foreach (ExportDWGSettings setting in collector)
                {
                    if (!string.IsNullOrEmpty(setting.Name))
                        setups.Add(setting.Name);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Error getting setups: {ex.Message}");
            }

            if (setups.Count == 0)
                setups.Add("(Default)");

            return setups;
        }

        public DWGExportOptions BuildExportOptions(ExportSettings settings)
        {
            DWGExportOptions options = null;

            if (!string.IsNullOrEmpty(settings.DwgExportSetupName)
                && settings.DwgExportSetupName != "(Default)")
            {
                try
                {
                    var collector = new FilteredElementCollector(_document)
                        .OfClass(typeof(ExportDWGSettings))
                        .Cast<ExportDWGSettings>()
                        .FirstOrDefault(s => s.Name == settings.DwgExportSetupName);

                    if (collector != null)
                        options = collector.GetDWGExportOptions();
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] Error loading setup: {ex.Message}");
                }
            }

            if (options == null)
                options = new DWGExportOptions();

            ConfigureCleanExportOptions(options, settings);
            return options;
        }

        private void ConfigureCleanExportOptions(DWGExportOptions options, ExportSettings settings)
        {
            TrySetProperty(options, "ExportingAreas", false);
            // Revit exports sheet views as XREF DWGs when MergedViews is false.
            // Combined output requires each sheet export to be self-contained.
            TrySetProperty(options, "MergedViews", true);
            options.SharedCoords = false;
            TrySetProperty(options, "ExportRoomsAndAreas", false);
            TrySetProperty(options, "PropOverrides", false);
            options.ExportOfSolids = SolidGeometry.Polymesh;

            var acaPrefType = typeof(DWGExportOptions).Assembly
                .GetTypes()
                .FirstOrDefault(t => t.Name == "ACAObjectPreference");
            if (acaPrefType != null)
                TrySetProperty(options, "ACAPreference", Enum.Parse(acaPrefType, "Geometry"));

            try
            {
                TrySetProperty(options, "TargetUnit", Enum.Parse(typeof(ExportUnit), "Millimeter"));
            }
            catch
            {
                TrySetProperty(options, "TargetUnit", ExportUnit.Default);
            }

            TrySetProperty(options, "Colors", GetEnumValue("ExportColorMode", "IndexColors"));
            TrySetProperty(options, "LineScaling", GetEnumValue("LineScaling", "ViewScale"));
            TrySetProperty(options, "HideReferencePlane", true);
            TrySetProperty(options, "HideScopeBox", true);
            TrySetProperty(options, "HideUnreferenceViewTags", true);
            TrySetProperty(options, "PreserveCoincidentLines", settings.PreserveCoincidentLines);

            options.FileVersion = GetAcadVersion(settings.DwgVersion);
        }

        private static void TrySetProperty(DWGExportOptions options, string propertyName, object value)
        {
            try
            {
                var property = typeof(DWGExportOptions).GetProperty(propertyName);
                if (property != null && property.CanWrite)
                    property.SetValue(options, value);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Failed to set {propertyName}: {ex.Message}");
            }
        }

        private static object GetEnumValue(string enumTypeName, string valueName)
        {
            try
            {
                var enumType = typeof(DWGExportOptions).Assembly
                    .GetTypes()
                    .FirstOrDefault(t => t.Name == enumTypeName && t.IsEnum);

                if (enumType != null)
                    return Enum.Parse(enumType, valueName);
            }
            catch { }

            return null;
        }

        public ExportResult ExportSheetsIndividually(
            List<SheetInfo> sheets, ExportSettings settings, DWGExportOptions options,
            IProgress<ExportProgressInfo> progress = null, CancellationToken cancellationToken = default)
        {
            var result = new ExportResult();
            var totalTimer = Stopwatch.StartNew();
            SmartScaleService smartScaleService = null;

            if (settings.SmartViewScale)
            {
                try
                {
                    smartScaleService = new SmartScaleService(_document);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] SmartScale init failed, continuing without it: {ex.Message}");
                }
            }

            try
            {
                EnsureOutputFolder(settings.OutputFolder);
                PrepareOutputFolderForExport(settings.OutputFolder);

                for (int i = 0; i < sheets.Count; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        Trace.WriteLine("[DwgExport] Export cancelled by user");
                        CleanupAfterCancel(result.ExportedFiles);
                        break;
                    }

                    var sheet = sheets[i];
                    var viewSheet = _document.GetElement(sheet.ElementId) as ViewSheet;
                    if (viewSheet == null)
                    {
                        result.SkippedSheets.Add(sheet.SheetNumber);
                        continue;
                    }

                    try
                    {
                        progress?.Report(new ExportProgressInfo
                        {
                            Phase = "Exporting",
                            CurrentItem = $"{sheet.SheetNumber} - {sheet.SheetName}",
                            Current = i + 1,
                            Total = sheets.Count
                        });

                        DispatcherDoEvents();

                        var sheetTimer = Stopwatch.StartNew();

                        if (smartScaleService != null)
                        {
                            using (var trans = new Transaction(_document, "Apply Smart Scale"))
                            {
                                try
                                {
                                    trans.Start();
                                    smartScaleService.ApplySmartScale(viewSheet, trans);
                                    trans.Commit();
                                }
                                catch
                                {
                                    if (trans.HasStarted())
                                        trans.RollBack();
                                    throw;
                                }
                            }
                        }

                        var filePath = ExportSingleSheet(viewSheet, sheet, settings, options);
                        sheetTimer.Stop();

                        if (smartScaleService != null)
                        {
                            using (var trans = new Transaction(_document, "Restore Scale"))
                            {
                                trans.Start();
                                smartScaleService.RestoreOriginalScale(viewSheet, trans);
                                trans.Commit();
                            }
                        }

                        if (!string.IsNullOrEmpty(filePath))
                        {
                            result.ExportedFiles.Add(filePath);
                            result.ExportedSheets.Add(sheet);
                            Trace.WriteLine($"[DwgExport] {sheet.SheetNumber} exported in {sheetTimer.ElapsedMilliseconds}ms");
                        }
                        else
                        {
                            result.FailedSheets.Add(sheet.SheetNumber);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (smartScaleService != null)
                        {
                            try
                            {
                                using (var trans = new Transaction(_document, "Restore Scale"))
                                {
                                    trans.Start();
                                    smartScaleService.RestoreOriginalScale(viewSheet, trans);
                                    trans.Commit();
                                }
                            }
                            catch (Exception innerEx)
                            {
                                Trace.WriteLine($"[DwgExport] Failed to restore scale: {innerEx.Message}");
                            }
                        }
                        result.FailedSheets.Add(sheet.SheetNumber);
                        Trace.WriteLine($"[DwgExport] Failed: {sheet.SheetNumber}: {ex.Message}");
                    }
                }
            }
            finally
            {
                if (smartScaleService != null)
                {
                    smartScaleService.ClearState();
                }
            }

            totalTimer.Stop();
            Trace.WriteLine($"[DwgExport] Total: {totalTimer.ElapsedMilliseconds}ms for {result.ExportedFiles.Count} sheets");

            if (result.FailedSheets.Count > 0)
                Trace.WriteLine($"[DwgExport] Failed sheets: {string.Join(", ", result.FailedSheets)}");

            return result;
        }

        private string ExportSingleSheet(ViewSheet viewSheet, SheetInfo sheetInfo, ExportSettings settings, DWGExportOptions options)
        {
            if (viewSheet == null) throw new ArgumentNullException(nameof(viewSheet));
            if (sheetInfo == null) throw new ArgumentNullException(nameof(sheetInfo));
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(settings.OutputFolder))
                throw new InvalidOperationException("OutputFolder is empty.");

            string fileName;
            try
            {
                fileName = GenerateFileName(sheetInfo, settings.FileNameTemplate, _document);
                fileName = SanitizeFileNamePart(fileName);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] File-name generation failed for '{sheetInfo.SheetNumber}': {ex}");
                throw;
            }

            var fullPath = Path.Combine(settings.OutputFolder, fileName + ".dwg");
            DeleteExportOutputIfExists(fullPath);

            try
            {
                ICollection<ElementId> sheetOnly = new List<ElementId> { viewSheet.Id };
                Trace.WriteLine($"[DwgExport] Exporting {viewSheet.SheetNumber} to {fullPath}");

                bool success = _document.Export(settings.OutputFolder, fileName, sheetOnly, options);

                if (success && File.Exists(fullPath))
                {
                    var fi = new FileInfo(fullPath);
                    Trace.WriteLine($"[DwgExport] OK: {fileName}.dwg ({fi.Length / 1024} KB)");
                    if (fi.Length < 1024)
                        Trace.WriteLine($"[DwgExport] WARNING: very small file ({fi.Length} bytes)");
                    if (DwgCleanupService.HasXRefFiles(fullPath))
                    {
                        int deleted = DwgCleanupService.CleanupXRefFiles(fullPath);
                        Trace.WriteLine($"[DwgExport] Cleaned up {deleted} XREF companion files for {fileName}.dwg");
                    }
                    return fullPath;
                }

                if (success && !File.Exists(fullPath))
                {
                    Trace.WriteLine($"[DwgExport] WARNING: success but file not found at {fullPath}");
                    var possibleFiles = Directory.GetFiles(settings.OutputFolder, fileName + "*.dwg");
                    if (possibleFiles.Length > 0)
                        return possibleFiles[0];
                }

                Trace.WriteLine($"[DwgExport] FAILED: {sheetInfo.SheetNumber} - export returned {success}");
                return null;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Exception exporting {sheetInfo.SheetNumber}: {ex}");
                return null;
            }
        }

        private static string SanitizeFileNamePart(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "Unnamed";

            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitized = new string(value
                .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
                .ToArray())
                .Trim();

            return string.IsNullOrWhiteSpace(sanitized) ? "Unnamed" : sanitized;
        }

        public static string GenerateFileName(SheetInfo sheet, string template, Document document = null)
        {
            if (sheet == null)
                return "";

            if (string.IsNullOrWhiteSpace(template))
                template = "{SheetNumber} - {SheetName}";

            string fileName = template
                .Replace("{SheetNumber}", sheet.SheetNumber ?? "")
                .Replace("{SheetName}", sheet.SheetName ?? "")
                .Replace("{PaperSize}", sheet.PaperSize ?? "")
                .Replace("{ProjectNumber}", GetProjectInfoValue(document, "Number"))
                .Replace("{ProjectName}", GetProjectInfoValue(document, "Name"))
                .Replace("{ProjectLocation}", GetProjectInfoValue(document, "PlaceName"))
                .Replace("{Author}", GetProjectInfoValue(document, "Author"))
                .Replace("{ClientName}", GetProjectInfoValue(document, "ClientName"))
                .Replace("{Date}", DateTime.Now.ToString("yyyyMMdd"))
                .Replace("{Time}", DateTime.Now.ToString("HHmm"));

            foreach (char c in Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '-');

            fileName = fileName.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
                fileName = !string.IsNullOrWhiteSpace(sheet.SheetNumber) ? sheet.SheetNumber : "Sheet";

            return fileName;
        }

        private static string GetProjectInfoValue(Document document, string propertyName)
        {
            try
            {
                var projectInfo = document?.ProjectInformation;
                if (projectInfo == null)
                    return "";

                var property = projectInfo.GetType().GetProperty(propertyName);
                var value = property == null ? null : property.GetValue(projectInfo, null);
                return value == null ? "" : value.ToString();
            }
            catch
            {
                return "";
            }
        }

        private void EnsureOutputFolder(string folder)
        {
            if (!Directory.Exists(folder))
            {
                try
                {
                    Directory.CreateDirectory(folder);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] Failed to create folder: {ex.Message}");
                }
            }
        }

        private void PrepareOutputFolderForExport(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return;

            var prefix = GetRevitGeneratedAssetPrefix();
            if (string.IsNullOrWhiteSpace(prefix))
                return;

            foreach (var file in Directory.EnumerateFiles(folder))
            {
                var name = Path.GetFileName(file);
                if (IsRevitGeneratedSupportFile(name, prefix))
                    DeleteExportOutputIfExists(file);
            }
        }

        private string GetRevitGeneratedAssetPrefix()
        {
            var title = _document?.Title;
            if (string.IsNullOrWhiteSpace(title))
                return "";

            return new string(title.Where(char.IsLetterOrDigit).ToArray());
        }

        private static bool IsRevitGeneratedSupportFile(string fileName, string prefix)
        {
            if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(prefix))
                return false;

            if (!fileName.StartsWith(prefix + "-", StringComparison.OrdinalIgnoreCase))
                return false;

            var ext = Path.GetExtension(fileName);
            return ext.Equals(".png", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".tif", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".tiff", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".bmp", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        }

        private static void DeleteExportOutputIfExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
            }
            catch (Exception ex)
            {
                var message =
                    "Cannot overwrite an existing export file because it is open or locked. " +
                    $"Close AutoCAD/AcCoreConsole or the file, then export again: {path}";
                Trace.WriteLine($"[DwgExport] {message}. {ex.Message}");
                throw new IOException(message, ex);
            }
        }

        private static void DispatcherDoEvents()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null) return;
            if (dispatcher.CheckAccess())
            {
                dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { }));
                dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { }));
            }
        }

        private void CleanupAfterCancel(List<string> files)
        {
            foreach (var file in files)
            {
                try { File.Delete(file); }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] Failed to clean up {file}: {ex.Message}");
                }
            }
        }

        private static ACADVersion GetAcadVersion(string version)
        {
            switch (version?.ToLower())
            {
                case "2025": 
                case "2024": 
                case "2023": 
                case "2022": 
                case "2021": 
                case "2020": return ACADVersion.R2018;
                case "2018": return ACADVersion.R2018;
                case "2013": return ACADVersion.R2013;
                case "2010": return ACADVersion.R2010;
                case "2007": return ACADVersion.R2007;
                default: return ACADVersion.R2018;
            }
        }
    }

    public class ExportProgressInfo
    {
        public string Phase { get; set; }
        public string CurrentItem { get; set; }
        public int Current { get; set; }
        public int Total { get; set; }
        public double Percentage => Total > 0 ? (double)Current / Total * 100 : 0;
    }

    public class DirectProgress<T> : IProgress<T>
    {
        private readonly Action<T> _report;

        public DirectProgress(Action<T> report)
        {
            _report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public void Report(T value)
        {
            _report(value);
        }
    }
}
