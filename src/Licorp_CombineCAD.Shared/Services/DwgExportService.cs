using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Helpers;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class DwgExportService
    {
        private readonly Document _document;
        private List<ElementId> _unloadedLinkIds;

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
                    {
                        options = collector.GetDWGExportOptions();
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] Error loading setup: {ex.Message}");
                }
            }

            if (options == null)
            {
                options = CreateSheetOnlyDWGOptions(settings);
            }
            else
            {
                OverrideOptionsFromSetup(options, settings);
            }

            return options;
        }

        private void OverrideOptionsFromSetup(DWGExportOptions options, ExportSettings settings)
        {
            try
            {
                Trace.WriteLine("[DwgExport] Overriding loaded setup options with clean export settings");
                ConfigureCleanExportOptions(options, settings);
                Trace.WriteLine("[DwgExport] Setup options overridden - XREF prevention applied");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Error overriding setup options: {ex.Message}");
            }
        }

        private DWGExportOptions CreateSheetOnlyDWGOptions(ExportSettings settings)
        {
            var options = new DWGExportOptions();
            Trace.WriteLine("[DwgExport] Creating ULTRA CLEAN options (Export+ method)");

            ConfigureCleanExportOptions(options, settings);

            Trace.WriteLine($"[DwgExport] FileVersion = {options.FileVersion}");
            Trace.WriteLine("[DwgExport] ================================");
            Trace.WriteLine("[DwgExport] ULTRA CLEAN EXPORT configured!");
            Trace.WriteLine("[DwgExport] Should export FULL GEOMETRY into 1 file");
            Trace.WriteLine("[DwgExport] ================================");

            return options;
        }

        private void TrySetProperty(DWGExportOptions options, string propertyName, object value)
        {
            try
            {
                var property = typeof(DWGExportOptions).GetProperty(propertyName);
                if (property != null && property.CanWrite)
                {
                    property.SetValue(options, value);
                    Trace.WriteLine($"[DwgExport] Set {propertyName} = {value}");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Failed to set {propertyName}: {ex.Message}");
            }
        }

        private void ConfigureCleanExportOptions(DWGExportOptions options, ExportSettings settings)
        {
            TrySetProperty(options, "ExportingAreas", false);
            // Revit exports sheet views as XREF DWGs when MergedViews is false.
            // Auto-bind means each sheet should become one self-contained DWG.
            var mergedViews = true;
            TrySetProperty(options, "MergedViews", mergedViews);
            Trace.WriteLine($"[DwgExport] MergedViews = {mergedViews} (self-contained sheet DWG for AutoCAD merge)");
            options.SharedCoords = false;
            TrySetProperty(options, "ExportRoomsAndAreas", false);
            TrySetProperty(options, "PropOverrides", false);
            options.ExportOfSolids = SolidGeometry.Polymesh;

            var acaPrefType = typeof(DWGExportOptions).Assembly
                .GetTypes()
                .FirstOrDefault(t => t.Name == "ACAObjectPreference");
            if (acaPrefType != null)
            {
                var geometryValue = Enum.Parse(acaPrefType, "Geometry");
                TrySetProperty(options, "ACAPreference", geometryValue);
            }

            try
            {
                var targetUnit = Enum.Parse(typeof(ExportUnit), "Millimeter");
                TrySetProperty(options, "TargetUnit", targetUnit);
                Trace.WriteLine("[DwgExport] TargetUnit = Millimeter");
            }
            catch
            {
                TrySetProperty(options, "TargetUnit", ExportUnit.Default);
                Trace.WriteLine("[DwgExport] TargetUnit = Default (fallback)");
            }

            TrySetProperty(options, "Colors", GetEnumValue("ExportColorMode", "IndexColors"));
            TrySetProperty(options, "LineScaling", GetEnumValue("LineScaling", "ViewScale"));

            TrySetProperty(options, "HideReferencePlane", true);
            TrySetProperty(options, "HideScopeBox", true);
            TrySetProperty(options, "HideUnreferenceViewTags", true);
            TrySetProperty(options, "PreserveCoincidentLines", settings.PreserveCoincidentLines);

            options.FileVersion = GetAcadVersion(settings.DwgVersion);
        }

        public ExportResult ExportSheetsIndividually(
            List<SheetInfo> sheets, ExportSettings settings, DWGExportOptions options,
            IProgress<ExportProgressInfo> progress = null, CancellationToken cancellationToken = default)
        {
            var result = new ExportResult();
            var totalTimer = System.Diagnostics.Stopwatch.StartNew();
            SmartScaleService smartScaleService = null;

            if (settings.SmartViewScale)
            {
                smartScaleService = new SmartScaleService(_document);
                Trace.WriteLine("[DwgExport] SmartScale enabled");
            }

            EnsureOutputFolder(settings.OutputFolder);

            if (settings.AutoBindXRef)
            {
                Trace.WriteLine("[DwgExport] AutoBindXRef enabled: using MergedViews=true; linked models remain loaded for complete sheet export.");
            }

            try
            {
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

                    if (sheet.HasNoView)
                    {
                        Trace.WriteLine($"[DwgExport] {sheet.SheetNumber} has no model viewport; exporting sheet content anyway.");
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

                        var sheetTimer = System.Diagnostics.Stopwatch.StartNew();

                        if (smartScaleService != null)
                        {
                            using (var trans = new Transaction(_document, "Apply Smart Scale"))
                            {
                                trans.Start();
                                smartScaleService.ApplySmartScale(viewSheet, trans);
                                trans.Commit();
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
if (_unloadedLinkIds != null && _unloadedLinkIds.Count > 0)
            {
                ReloadLinkedModels();
                _unloadedLinkIds = null;
            }
        }

        totalTimer.Stop();
        Trace.WriteLine($"[DwgExport] Total export time: {totalTimer.ElapsedMilliseconds}ms for {result.ExportedFiles.Count} sheets");

        if (result.FailedSheets.Count > 0)
        {
            Trace.WriteLine($"[DwgExport] Failed sheets: {string.Join(", ", result.FailedSheets)}");
        }

        return result;
    }

        private string ExportSingleSheet(ViewSheet viewSheet, SheetInfo sheetInfo, ExportSettings settings, DWGExportOptions options)
        {
            string fileName = GenerateFileName(sheetInfo, settings.FileNameTemplate, _document);
            string fullPath = Path.Combine(settings.OutputFolder, fileName + ".dwg");

            if (File.Exists(fullPath))
            {
                try { File.Delete(fullPath); }
                catch (Exception delEx)
                {
                    Trace.WriteLine($"[DwgExport] Failed to delete existing file: {delEx.Message}");
                }
            }

            try
            {
                ICollection<ElementId> sheetOnly = new List<ElementId> { viewSheet.Id };
                Trace.WriteLine($"[DwgExport] Exporting sheet {viewSheet.SheetNumber} to {fullPath}");

                bool success = _document.Export(settings.OutputFolder, fileName, sheetOnly, options);

                if (success && File.Exists(fullPath))
                {
                    var fi = new FileInfo(fullPath);
                    Trace.WriteLine($"[DwgExport] SUCCESS: {fileName}.dwg ({fi.Length / 1024} KB)");

                    if (fi.Length < 1024)
                    {
                        Trace.WriteLine($"[DwgExport] WARNING: File is very small ({fi.Length} bytes) - may be empty!");
                    }

                    if (DwgCleanupService.HasXRefFiles(fullPath))
                    {
                        Trace.WriteLine($"[DwgExport] XREF companion files detected - Revit split the export");
                    }

                    return fullPath;
                }

                if (success && !File.Exists(fullPath))
                {
                    Trace.WriteLine($"[DwgExport] WARNING: Export returned success but file not found at {fullPath}");
                    var possibleFiles = Directory.GetFiles(settings.OutputFolder, fileName + "*.dwg");
                    if (possibleFiles.Length > 0)
                    {
                        Trace.WriteLine($"[DwgExport] Found alternative: {possibleFiles[0]}");
                        return possibleFiles[0];
                    }
                }

                Trace.WriteLine($"[DwgExport] FAILED: {sheetInfo.SheetNumber} - export returned {success}");
                return null;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Exception exporting {sheetInfo.SheetNumber}: {ex.Message}");
                return null;
            }
        }

        private List<ElementId> UnloadLinkedModels()
        {
            var unloadedIds = new List<ElementId>();

            try
            {
                var linkTypes = new FilteredElementCollector(_document)
                    .OfClass(typeof(RevitLinkType))
                    .Cast<RevitLinkType>()
                    .Where(lt => lt.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
                    .ToList();

                if (linkTypes.Count == 0)
                    return unloadedIds;

                using (var trans = new Transaction(_document, "Unload Links for DWG Export"))
                {
                    trans.Start();

                    foreach (var linkType in linkTypes)
                    {
                        try
                        {
                            linkType.Unload(null);
                            unloadedIds.Add(linkType.Id);
                            Trace.WriteLine($"[DwgExport] Unloaded: {linkType.Name}");
                        }
                        catch (Exception ex)
                        {
                            Trace.WriteLine($"[DwgExport] Failed to unload {linkType.Name}: {ex.Message}");
                        }
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Error unloading links: {ex.Message}");
            }

            return unloadedIds;
        }

        private void ReloadLinkedModels()
        {
            if (_unloadedLinkIds == null || _unloadedLinkIds.Count == 0)
                return;

            try
            {
                using (var trans = new Transaction(_document, "Reload Links after DWG Export"))
                {
                    trans.Start();

                    foreach (var linkId in _unloadedLinkIds)
                    {
                        var linkType = _document.GetElement(linkId) as RevitLinkType;
                        if (linkType != null)
                        {
                            try
                            {
                                linkType.Reload();
                                Trace.WriteLine($"[DwgExport] Reloaded: {linkType.Name}");
                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"[DwgExport] Failed to reload {linkType.Name}: {ex.Message}");
                            }
                        }
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[DwgExport] Error reloading links: {ex.Message}");
            }
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
                .Replace("{ProjectName}", GetProjectInfoValue(document, "Name"));

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '-');
            }

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
                    Trace.WriteLine($"[DwgExport] Created folder: {folder}");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] Failed to create folder: {ex.Message}");
                }
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
                try
                {
                    File.Delete(file);
                    Trace.WriteLine($"[DwgExport] Cleaned up: {file}");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"[DwgExport] Failed to clean up {file}: {ex.Message}");
                }
            }
        }

        private ACADVersion GetAcadVersion(string version)
        {
            switch (version?.ToLower())
            {
                case "2018": return ACADVersion.R2018;
                case "2013": return ACADVersion.R2013;
                case "2010": return ACADVersion.R2010;
                case "2007": return ACADVersion.R2007;
                default: return ACADVersion.R2018;
            }
        }

        private object GetEnumValue(string enumTypeName, string valueName)
        {
            try
            {
                var enumType = typeof(DWGExportOptions).Assembly
                    .GetTypes()
                    .FirstOrDefault(t => t.Name == enumTypeName && t.IsEnum);

                if (enumType != null)
                {
                    return Enum.Parse(enumType, valueName);
                }
            }
            catch { }

            return null;
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
