using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Diagnostics;
using Newtonsoft.Json;

namespace Licorp_MergeSheets
{
    public class LayoutMerger
    {
        private const double LayoutSpacing = 50.0;
        private const string PaperBackgroundLayerName = "LICORP_PAPER_BACKGROUND";
        private const string ViewportLayerName = "VIEWPORTS";
        private const double PaperBackgroundFallbackWidth = 1066.8;
        private const double PaperBackgroundFallbackHeight = 762.0;
        private const double ModelSpaceSheetMinGap = 25.0;
        private const int AutoCadLayoutNameMaxLength = 31;
        
        // Track ModelSpace offset for each source file
        private Dictionary<string, Vector3d> _msOffsets = new Dictionary<string, Vector3d>();
        private double _currentMsXOffset = 0.0;

        // Cache for geometry extents to avoid recalculating
        private Dictionary<string, ExtentsCacheEntry> _extentsCache = new Dictionary<string, ExtentsCacheEntry>(StringComparer.OrdinalIgnoreCase);

        // Cache for PlotSettings to avoid repeated RefreshLists calls
        private static string _cachedPlotDevice = null;

        private void ReportProgress(MergeConfig config, int current, int total, string layoutName)
        {
            try
            {
                config?.ProgressCallback?.Invoke(current, total, layoutName);

                if (!string.IsNullOrEmpty(config?.StatusPath))
                {
                    var progressData = new
                    {
                        Phase = "Merging",
                        Current = current,
                        Total = total,
                        CurrentItem = layoutName,
                        Percentage = total > 0 ? (double)current / total * 100 : 0,
                        Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    };
                    var dir = Path.GetDirectoryName(config.StatusPath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    File.WriteAllText(config.StatusPath + ".progress.json",
                        JsonConvert.SerializeObject(progressData, Formatting.Indented));
                }
            }
            catch { }
        }
        private static List<string> _cachedCanonicalMedia = null;
        private static DateTime _cachedMediaAt = DateTime.MinValue;
        private static readonly TimeSpan MediaCacheExpiry = TimeSpan.FromMinutes(5);

        private void EnsurePlotSettingsRefreshed(Layout layout)
        {
            string currentDevice = null;
            try { currentDevice = layout.PlotConfigurationName; } catch { }

            if (string.IsNullOrEmpty(currentDevice))
                return;

            bool needsRefresh = _cachedCanonicalMedia == null ||
                !string.Equals(_cachedPlotDevice, currentDevice, StringComparison.OrdinalIgnoreCase);

            if (needsRefresh)
            {
                var psv = PlotSettingsValidator.Current;
                psv.RefreshLists(layout);
                try
                {
                    var mediaNames = psv.GetCanonicalMediaNameList(layout);
                    _cachedCanonicalMedia = new List<string>();
                    foreach (string name in mediaNames)
                    {
                        _cachedCanonicalMedia.Add(name);
                    }
                    _cachedPlotDevice = currentDevice;
                    _cachedMediaAt = DateTime.Now;
                    AcadLogger.LogInfo($"PlotSettings cache: refreshed for device '{currentDevice}', media count={_cachedCanonicalMedia.Count}");
                }
                catch (Exception ex)
                {
                    AcadLogger.LogWarning($"PlotSettings cache: failed to get media list: {ex.Message}");
                }
            }
        }

        private void ClearPlotSettingsCache()
        {
            _cachedPlotDevice = null;
            _cachedCanonicalMedia = null;
            _cachedMediaAt = DateTime.MinValue;
            _mediaSizeCache.Clear();
            _correctedPaperCache.Clear();
        }

        private class ExtentsCacheEntry
        {
            public Extents3d Extents { get; set; }
            public int EntityCount { get; set; }
            public int ExtentsEntityCount { get; set; }
            public DateTime CachedAt { get; set; }
            public bool IsModelSpace { get; set; }
        }

        private string GetExtentsCacheKey(string filePath, string layoutName)
        {
            if (string.IsNullOrEmpty(filePath))
                return layoutName ?? string.Empty;
            return (filePath + "|" + (layoutName ?? "Model")).ToLower();
        }

        private bool TryGetExtentsCache(string key, out ExtentsCacheEntry cache)
        {
            return _extentsCache.TryGetValue(key, out cache);
        }

        private void SetExtentsCache(string key, ExtentsCacheEntry cache)
        {
            _extentsCache[key] = cache;
        }

        private void ClearExtentsCache()
        {
            _extentsCache.Clear();
        }

        public bool MergeToMultiLayout(MergeConfig config)
        {
            try
            {
                AcadLogger.LogSection("MergeToMultiLayout");
                AcadLogger.LogInfo($"Output path: {config.OutputPath}");
                AcadLogger.LogInfo($"Source files: {config.SourceFiles?.Count ?? 0}");

                if (config.SourceFiles == null || config.SourceFiles.Count == 0)
                {
                    AcadLogger.LogError("No source files provided");
                    return false;
                }

                string baseFile = null;
                foreach (var sf in config.SourceFiles)
                {
                    if (File.Exists(sf.Path))
                    {
                        baseFile = sf.Path;
                        break;
                    }
                }

                if (baseFile == null)
                {
                    AcadLogger.LogError("No valid source files found");
                    return false;
                }

                AcadLogger.LogInfo($"Using base file: {baseFile}");

                var outputDb = new Database(false, true);
                _msOffsets.Clear();
                _currentMsXOffset = 0.0;

                bool keepViewportLive =
                    !string.Equals(config?.ViewportMode, "Baked", StringComparison.OrdinalIgnoreCase);

                AcadLogger.LogInfo($"Viewport handling: mode={(keepViewportLive ? "Live" : "Baked")}");

                var sourceInfos = new List<SourceFileInfo>();
                var pendingScheduleOnlyLayouts = new List<SourceFileInfo>();
                var usedLayoutNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                using (outputDb)
                {
                    outputDb.ReadDwgFile(baseFile, FileShare.ReadWrite, true, "");
                    outputDb.CloseInput(true);
                    AcadLogger.LogInfo("Base file opened successfully");
                    BindXrefsSafe(outputDb);

                    using (var trans = outputDb.TransactionManager.StartTransaction())
                    {
                        var firstSource = config.SourceFiles.First(s => File.Exists(s.Path));
                        var firstLayoutName = GetSafeAutoCadLayoutName(firstSource.Layout ?? "Layout1", usedLayoutNames);

                        if (!string.IsNullOrEmpty(firstSource.Layout))
                        {
                            RenameLayoutInDb(outputDb, "Layout1", firstLayoutName);
                            AcadLogger.LogInfo($"Renamed base layout to '{firstLayoutName}'");
                        }

                        var baseStats = GetModelSpaceStats(outputDb, trans);
                        LogModelSpaceStats("Base after bind", baseStats);

                        var baseExtents = baseStats.Extents;
                        double baseWidth = baseExtents.MaxPoint.X - baseExtents.MinPoint.X;
                        if (baseWidth <= 0)
                        {
                            AcadLogger.LogWarning($"Base MS extents invalid (width={baseWidth}), using safe default");
                            baseWidth = 100000;
                        }

                        var firstLayout = GetSourceLayout(outputDb, trans, firstLayoutName);
                        var baseOccupiedExtents = baseExtents;

                        if (firstLayout != null)
                        {
                            baseOccupiedExtents = CombineExtents(
                                baseExtents,
                                GetLayoutModelViewExtents(outputDb, trans, firstLayout, baseExtents, "Base layout"));

                            var firstBtr = (BlockTableRecord)trans.GetObject(firstLayout.BlockTableRecordId, OpenMode.ForWrite);
                            var firstViewports = CollectModelViewportInfos(trans, firstBtr, $"BASE usable viewports: {firstLayoutName}");

                            if (keepViewportLive)
                            {
                                AcadLogger.LogInfo($"BASE: Keep live viewport mode enabled; skip paper bake/erase for '{firstLayoutName}'");
                            }
                            else
                            {
                                int baseBakedCount = BakeModelViewsToPaperSpace(
                                    outputDb, trans, outputDb, trans, firstBtr, firstViewports, firstLayoutName);

                                int baseErasedViewportCount = baseBakedCount > 0
                                    ? EraseAllLayoutViewports(trans, firstBtr, firstLayoutName)
                                    : 0;

                                AcadLogger.LogInfo(
                                    $"BASE: Baked {baseBakedCount} model entity clone(s) to PaperSpace and erased " +
                                    $"{baseErasedViewportCount} viewport(s) for '{firstLayoutName}'");
                            }

                            var paperCtx = EnsureLayoutPaperContextFromGeometry(
                                outputDb, trans, firstLayout, firstBtr, firstLayoutName, "BASE", firstSource.PaperSize);

                            if (keepViewportLive)
                            {
                                TransformExistingViewportsForPaperRotation(
                                    trans, firstBtr, paperCtx, firstLayoutName);
                            }

                            var ps = new PlotSettings(firstLayout.ModelType);
                            ps.CopyFrom(firstLayout);

                            sourceInfos.Add(new SourceFileInfo
                            {
                                FilePath = baseFile,
                                LayoutName = firstLayoutName,
                                MsOffset = new Vector3d(0, 0, 0),
                                MsExtents = baseExtents,
                                ModelType = firstLayout.ModelType,
                                PlotSettings = ps
                            });
                        }

                        _currentMsXOffset = baseOccupiedExtents.MaxPoint.X + GetLayoutGap(baseOccupiedExtents);
                        AcadLogger.LogInfo($"Base occupied max X: {baseOccupiedExtents.MaxPoint.X:F2}, next visible min X: {_currentMsXOffset:F2}");

                        trans.Commit();
                    }

                    int clonedCount = 1;
                    int fileIndex = 2;
                    var failedLayouts = new List<string>();
                    int totalToProcess = config.SourceFiles.Count(s => File.Exists(s.Path) && !s.Path.Equals(baseFile, StringComparison.OrdinalIgnoreCase));
                    int processedCount = 0;

                    foreach (var source in config.SourceFiles)
                    {
                        if (!File.Exists(source.Path))
                        {
                            AcadLogger.LogWarning($"Source not found: {source.Path}");
                            continue;
                        }

                        if (source.Path.Equals(baseFile, StringComparison.OrdinalIgnoreCase))
                            continue;

                        string requestedName = source.Layout ?? $"Layout{fileIndex}";
                        string desiredName = GetSafeAutoCadLayoutName(requestedName, usedLayoutNames);
                        AcadLogger.LogInfo($"Processing: {Path.GetFileName(source.Path)} -> Layout '{desiredName}'");

                        PlotSettings savedPlotSettings = null;
                        bool modelType = false;
                        Vector3d msOffset = Vector3d.ZAxis;
                        Extents3d srcExtents = new Extents3d();

                        try
                        {

                        var sourceDb = new Database(false, true);
                        using (sourceDb)
                        {
                            sourceDb.ReadDwgFile(source.Path, FileShare.ReadWrite, true, "");
                            sourceDb.CloseInput(true);
                            BindXrefsSafe(sourceDb);

                            using (var srcTrans = sourceDb.TransactionManager.StartTransaction())
                            using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                            {
                                RenameBlocksInDb(sourceDb, srcTrans, $"File{fileIndex}_");

                                // Merge layers from source if enabled
                                if (config.MergeLayers)
                                {
                                    int mergedLayers = MergeLayersFromSource(sourceDb, outputDb, srcTrans, outputTrans);
                                    AcadLogger.LogInfo($"Layer merge result for '{desiredName}': {mergedLayers} layer(s) added");
                                }

                                var srcMs = (BlockTableRecord)srcTrans.GetObject(
                                    SymbolUtilityServices.GetBlockModelSpaceId(sourceDb), OpenMode.ForRead);
                                var outputMs = (BlockTableRecord)outputTrans.GetObject(
                                    SymbolUtilityServices.GetBlockModelSpaceId(outputDb), OpenMode.ForWrite);

                                var srcStats = GetModelSpaceStats(sourceDb, srcTrans);
                                LogModelSpaceStats($"Source after bind: {Path.GetFileName(source.Path)}", srcStats);

                                var srcLayout = GetSourceLayout(sourceDb, srcTrans, source.Layout);
                                if (srcLayout == null)
                                {
                                    AcadLogger.LogError($"No layout found in {Path.GetFileName(source.Path)}");
                                    fileIndex++;
                                    continue;
                                }

                                modelType = srcLayout.ModelType;
                                savedPlotSettings = new PlotSettings(srcLayout.ModelType);
                                savedPlotSettings.CopyFrom(srcLayout);

                                var viewportViewExtents = GetLayoutModelViewExtents(
                                    sourceDb, srcTrans, srcLayout, srcStats.Extents, desiredName);
                                var sourceVisibleExtents = CombineExtents(srcStats.Extents, viewportViewExtents);

                                AcadLogger.LogInfo($"GD1: {desiredName} viewport view extents {FormatExtents(viewportViewExtents)}");
                                AcadLogger.LogInfo($"GD1: {desiredName} source visible extents (combined) {FormatExtents(sourceVisibleExtents)}");

                                var msIds = new ObjectIdCollection();
                                foreach (ObjectId id in srcMs)
                                {
                                    try
                                    {
                                        var msEnt = srcTrans.GetObject(id, OpenMode.ForRead, false) as Entity;
                                        if (msEnt == null) continue;

                                        // In live-viewport mode, skip BlockReferences that originate from
                                        // xref / overlay blocks (e.g. Revit-exported sheets embedded as
                                        // a BlockReference in ModelSpace at viewport scale ~0.01).
                                        // Cloning them into the output DB makes them appear ON TOP of the
                                        // recreated Viewport entity so that clicking the drawing area selects
                                        // a BLOCK REFERENCE instead of the VIEWPORT.
                                        if (keepViewportLive && msEnt is BlockReference msBr)
                                        {
                                            bool isXrefOrOverlay = false;
                                            bool isAnonymous = false;
                                            bool isLargeScaleBR = false;
                                            try
                                            {
                                                var btrDef = srcTrans.GetObject(msBr.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                                if (btrDef != null)
                                                {
                                                    isXrefOrOverlay = btrDef.IsFromExternalReference || btrDef.IsFromOverlayReference;
                                                    isAnonymous = btrDef.IsAnonymous;
                                                }
                                                // Scale < 1 in all axes typically means a "sheet-embedded" block
                                                // (Revit exports use 1:96 = 0.0104 or similar sub-1 scales)
                                                double bScale = Math.Abs(msBr.ScaleFactors.X);
                                                isLargeScaleBR = bScale > 0.0 && bScale < 1.0;
                                            }
                                            catch { }

                                            if (isXrefOrOverlay || isAnonymous || isLargeScaleBR)
                                            {
                                                AcadLogger.LogInfo(
                                                    $"GD1: SKIP ModelSpace BlockRef '{msBr.Name}' " +
                                                    $"xref={isXrefOrOverlay} anon={isAnonymous} smallScale={isLargeScaleBR} " +
                                                    $"scale={msBr.ScaleFactors.X:F6} (would shadow viewport in output)");
                                                continue;
                                            }
                                        }

                                        msIds.Add(id);
                                    }
                                    catch { }
                                }

                                msOffset = new Vector3d(_currentMsXOffset - viewportViewExtents.MinPoint.X, 0, 0);
                                _msOffsets[desiredName] = msOffset;
                                AcadLogger.LogInfo($"GD1: {desiredName} msOffset={FormatVector(msOffset)}, target visible min X={_currentMsXOffset:F2}");

                                var msIdMap = new IdMapping();
                                if (msIds.Count > 0)
                                {
                                    sourceDb.WblockCloneObjects(
                                        msIds, outputMs.ObjectId, msIdMap, DuplicateRecordCloning.Replace, false);

                                    var transform = Matrix3d.Displacement(msOffset);
                                    foreach (ObjectId srcId in msIds)
                                    {
                                        if (!msIdMap.Contains(srcId)) continue;
                                        ObjectId destId = msIdMap[srcId].Value;
                                        if (destId.IsNull) continue;
                                        try
                                        {
                                            var ent = outputTrans.GetObject(destId, OpenMode.ForWrite) as Entity;
                                            if (ent != null)
                                                ent.TransformBy(transform);
                                        }
                                        catch { }
                                    }
                                }

                                srcExtents = srcStats.Extents;
                                _currentMsXOffset = msOffset.X + sourceVisibleExtents.MaxPoint.X + GetLayoutGap(sourceVisibleExtents);
                                AcadLogger.LogInfo($"GD1: Updated next visible min X to {_currentMsXOffset:F2}");

                                var destBtrId = CreateNewLayoutInDb(outputDb, outputTrans, desiredName);
                                if (destBtrId.IsNull)
                                    destBtrId = ReuseEmptyDefaultLayout(outputDb, outputTrans, desiredName);

                                if (destBtrId.IsNull)
                                {
                                    AcadLogger.LogError($"Failed to create layout '{desiredName}'. Deferring as schedule-only layout.");
                                    pendingScheduleOnlyLayouts.Add(new SourceFileInfo
                                    {
                                        FilePath = source.Path,
                                        LayoutName = desiredName,
                                        MsOffset = msOffset,
                                        MsExtents = srcExtents,
                                        ModelType = modelType,
                                        PlotSettings = savedPlotSettings
                                    });
                                    fileIndex++;
                                    continue;
                                }

                                var destBtr = (BlockTableRecord)outputTrans.GetObject(destBtrId, OpenMode.ForWrite);
                                var destLayout = (Layout)outputTrans.GetObject(destBtr.LayoutId, OpenMode.ForWrite);

                                var savedBtrId = destLayout.BlockTableRecordId;
                                var savedTabOrder = destLayout.TabOrder;
                                ApplyLayoutPlotSettingsSafely(
                                    destLayout, savedPlotSettings, desiredName, savedBtrId, savedTabOrder, "GD2");

                                var srcBtr = (BlockTableRecord)srcTrans.GetObject(srcLayout.BlockTableRecordId, OpenMode.ForRead);
                                var sourceViewports = CollectModelViewportInfos(srcTrans, srcBtr, $"SRC usable viewports: {desiredName}");

                                double srcMaxViewArea = 0.0;
                                foreach (ObjectId id in srcBtr)
                                {
                                    try
                                    {
                                        var v = srcTrans.GetObject(id, OpenMode.ForRead, false) as Viewport;
                                        if (IsRawModelViewport(v))
                                            srcMaxViewArea = Math.Max(srcMaxViewArea, GetViewportViewArea(v));
                                    }
                                    catch { }
                                }

                                var psIds = new ObjectIdCollection();
                                foreach (ObjectId id in srcBtr)
                                {
                                    try
                                    {
                                        var ent = srcTrans.GetObject(id, OpenMode.ForRead, false) as Entity;
                                        if (ent == null)
                                            continue;

                                        if (string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        if (keepViewportLive && ent is Viewport)
                                            continue;

                                        // In live viewport mode, avoid cloning PaperSpace block references.
                                        // They sit on top of viewports and capture clicks/double-clicks (BEDIT),
                                        // making the user select a BLOCK REFERENCE instead of the MVIEW viewport.
                                        // Note: ModelSpace BlockReferences are also filtered above (msIds loop)
                                        // to prevent the same issue for geometry visible through the viewport.
                                        // EXCEPTION: Title block references must be preserved.
                                        if (keepViewportLive && ent is BlockReference br && !IsTitleBlockReference(br, srcTrans))
                                            continue;

                                        var sourceVp = ent as Viewport;
                                        if (sourceVp != null && IsUtilityViewport(sourceVp, srcMaxViewArea))
                                            continue;

                                        psIds.Add(id);
                                    }
                                    catch { }
                                }

                                var psIdMap = new IdMapping();
                                if (psIds.Count > 0)
                                {
                                    sourceDb.WblockCloneObjects(
                                        psIds, destBtrId, psIdMap, DuplicateRecordCloning.Replace, false);
                                }

                                LogViewportCollection(outputTrans, destBtr, $"DEST after PS clone: {desiredName}");

                                if (keepViewportLive)
                                {
                                    int erasedViewportCount = EraseAllLayoutViewportsForRecreate(
                                        outputTrans, destBtr, desiredName);

                                    AcadLogger.LogInfo(
                                        $"GD2: Erased cloned viewports for '{desiredName}': {erasedViewportCount}");

                                    int vpCreatedCount = RecreateLayoutViewports(
                                        outputTrans, destBtr, sourceViewports, desiredName, msOffset);

                                    EnsureLayoutViewportsOnTop(outputTrans, destBtr, desiredName);

                                    AcadLogger.LogInfo(
                                        $"GD2: Keep live viewport mode enabled; recreated {vpCreatedCount} viewport(s) for '{desiredName}'");

                                    LogViewportCollection(outputTrans, destBtr, $"DEST after RECREATE: {desiredName}");

                                    LogLayoutDiag(desiredName, "ManualLayoutRecreate",
                                        sourceViewports.Count, vpCreatedCount, 0, erasedViewportCount, msOffset, true);
                                }
                                else
                                {
                                    int bakedCount = BakeModelViewsToPaperSpace(
                                        sourceDb, srcTrans, outputDb, outputTrans, destBtr, sourceViewports, desiredName);

                                    int erasedViewportCount = sourceViewports.Count > 0
                                        ? EraseAllLayoutViewports(outputTrans, destBtr, desiredName)
                                        : 0;

                                    AcadLogger.LogInfo(
                                        $"GD2: Baked {bakedCount} model entity clone(s) to PaperSpace and erased " +
                                        $"{erasedViewportCount} viewport(s) for '{desiredName}'");

                                    LogLayoutDiag(desiredName, "ManualLayoutClone",
                                        sourceViewports.Count, 0, bakedCount, erasedViewportCount, msOffset, false);
                                }

                                EnsureLayoutPaperContextFromGeometry(
                                    outputDb, outputTrans, destLayout, destBtr, desiredName, "GD2", source.PaperSize);

                                if (srcStats.EntityCount == 0 && !LayoutHasContent(destBtr, outputTrans))
                                    AddSchedulePlaceholderContent(destBtr, outputTrans, desiredName, destLayout);

                                sourceInfos.Add(new SourceFileInfo
                                {
                                    FilePath = source.Path,
                                    LayoutName = desiredName,
                                    MsOffset = msOffset,
                                    MsExtents = srcExtents,
                                    ModelType = modelType,
                                    PlotSettings = savedPlotSettings
                                });

                                outputTrans.Commit();
                                srcTrans.Commit();
                            }
                        }

                        clonedCount++;
                        processedCount++;
                        AcadLogger.LogInfo($"Successfully cloned layout '{desiredName}'");
                        ReportProgress(config, processedCount, totalToProcess, desiredName);
                        }
                        catch (System.Exception layoutEx)
                        {
                            failedLayouts.Add($"{desiredName}: {layoutEx.Message}");
                            AcadLogger.LogWarning($"FAILED layout '{desiredName}' from '{Path.GetFileName(source.Path)}': {layoutEx.Message}");
                        }
                        fileIndex++;
                    }

                    clonedCount += EnsurePendingScheduleOnlyLayouts(outputDb, pendingScheduleOnlyLayouts, sourceInfos);
                    AcadLogger.LogInfo($"Total layouts cloned: {clonedCount}");

                    if (failedLayouts.Count > 0)
                    {
                        AcadLogger.LogWarning($"ERROR RECOVERY: {failedLayouts.Count}/{totalToProcess} layout(s) failed:");
                        foreach (var failed in failedLayouts)
                            AcadLogger.LogWarning($"  - {failed}");
                    }

                    if (processedCount == 0 && clonedCount <= 1)
                    {
                        AcadLogger.LogError("All source layouts failed. No output generated.");
                        return false;
                    }

                    CleanupDefaultLayouts(outputDb);
                    RemovePaperBackgroundPresentation(outputDb, "MultiLayout");

                    RegenerateLayouts(outputDb, keepViewportLive ? "MultiLayout-Live" : "MultiLayout-Baked");

                    var dwgVersion = GetDwgVersion(config.DwgVersion);
                    outputDb.SaveAs(config.OutputPath, dwgVersion);
                    ClearExtentsCache();
                    ClearPlotSettingsCache();
                    AcadLogger.LogInfo($"Saved to: {config.OutputPath}");

                    var outputPathForRegen = config.OutputPath;
                    var layoutNamesForRegen = sourceInfos.Select(s => s.LayoutName).ToList();
                    AcadLogger.LogInfo(
                        $"POST-SAVE REGEN: scheduling deferred layout regen for {layoutNamesForRegen.Count} layout(s) via Application.Idle");
                    SchedulePostSaveRegen(outputPathForRegen, layoutNamesForRegen);

                    DisposePlotSettings(sourceInfos);
                    DisposePlotSettings(pendingScheduleOnlyLayouts);

                    AcadLogger.LogSection("MergeToMultiLayout Complete");
                    return true;
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"MultiLayout error: {ex.Message}");
                AcadLogger.LogError($"Stack: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Schedules a deferred regen of the output DWG via Application.Idle (fires after
        /// the current command stack unwinds). Opens the saved file, runs REGENALL to build
        /// viewport display lists, then QSAVE to persist them. This avoids the fatal crash
        /// caused by calling DocumentManager.Open synchronously inside a command.
        /// </summary>
        private void SchedulePostSaveRegen(string outputPath, List<string> layoutNames)
        {
            EventHandler idleHandler = null;
            idleHandler = (sender, e) =>
            {
                try
                {
                    Application.Idle -= idleHandler;

                    AcadLogger.LogInfo($"POST-SAVE REGEN [Idle]: opening '{outputPath}'");

                    var dm = Application.DocumentManager;
                    var regenDoc = dm.Open(outputPath, false);
                    if (regenDoc == null)
                    {
                        AcadLogger.LogWarning("POST-SAVE REGEN [Idle]: document open returned null");
                        return;
                    }

                    dm.MdiActiveDocument = regenDoc;

                    void Send(string cmd)
                    {
                        AcadLogger.LogInfo($"POST-SAVE REGEN [Idle]: send {cmd}");
                        regenDoc.SendStringToExecute(cmd + " ", true, false, false);
                    }

                    Send("_.TILEMODE 0");

                    int layoutCount = layoutNames?.Count ?? 0;
                    AcadLogger.LogInfo($"POST-SAVE REGEN [Idle]: queueing layout activation for {layoutCount} layout(s)");

                    if (layoutNames != null)
                    {
                        foreach (var layoutName in layoutNames)
                        {
                            if (string.IsNullOrWhiteSpace(layoutName))
                                continue;

                            AcadLogger.LogInfo($"POST-SAVE REGEN [Idle]: queue layout '{layoutName}'");
                            Send($"_.LAYOUT _Set \"{layoutName}\"");
                            Send("_.PSPACE");
                            Send("_.REGEN");
                        }
                    }

                    Send("_.REGENALL");
                    Send("_.QSAVE");

                    AcadLogger.LogInfo("POST-SAVE REGEN [Idle]: queued TILEMODE0 + layout-by-layout REGEN + REGENALL + QSAVE");
                }
                catch (System.Exception ex)
                {
                    try { Application.Idle -= idleHandler; } catch { }
                    AcadLogger.LogWarning($"POST-SAVE REGEN [Idle]: failed: {ex.Message}");
                }
            };

            Application.Idle += idleHandler;
            AcadLogger.LogInfo("POST-SAVE REGEN [Idle]: handler attached");
        }

        public bool MergeToSingleLayout(MergeConfig config)
        {
            try
            {
                AcadLogger.LogSection("MergeToSingleLayout");
                AcadLogger.LogInfo($"Output path: {config.OutputPath}");
                AcadLogger.LogInfo($"Source files: {config.SourceFiles?.Count ?? 0}");

                var outputDb = new Database(false, true);

                using (outputDb)
                {
                    var firstFile = config.SourceFiles.FirstOrDefault(f => File.Exists(f.Path));
                    if (firstFile == null)
                    {
                        AcadLogger.LogError("No valid source file for base database");
                        return false;
                    }

                    outputDb.ReadDwgFile(firstFile.Path, FileShare.ReadWrite, true, "");
                    BindXrefsSafe(outputDb);

                    var allSourceExtents = new List<Extents3d>();
                    var allClonedIds = new List<List<ObjectId>>();
                    var allSourceDbs = new List<Database>();

                    using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                    {
                        // Target Layout: Reuse "Layout1" or create "CombinedSheet"
                        string targetLayoutName = "CombinedSheet";
                        var layouts = (DBDictionary)outputTrans.GetObject(outputDb.LayoutDictionaryId, OpenMode.ForRead);
                        ObjectId targetLayoutBtrId = ObjectId.Null;

                        if (layouts.Contains("Layout1"))
                        {
                            var layout1 = (Layout)outputTrans.GetObject(layouts.GetAt("Layout1"), OpenMode.ForWrite);
                            layout1.LayoutName = targetLayoutName;
                            targetLayoutBtrId = layout1.BlockTableRecordId;
                            AcadLogger.LogInfo($"Renamed 'Layout1' to '{targetLayoutName}'");
                        }
                        else
                        {
                            // Fallback: Create new layout if Layout1 doesn't exist (rare)
                            var lm = LayoutManager.Current;
                            lm.CreateLayout(targetLayoutName);
                            var newLayout = (Layout)outputTrans.GetObject(layouts.GetAt(targetLayoutName), OpenMode.ForRead);
                            targetLayoutBtrId = newLayout.BlockTableRecordId;
                        }

                        var targetBtr = (BlockTableRecord)outputTrans.GetObject(targetLayoutBtrId, OpenMode.ForWrite);
                        AcadLogger.LogInfo($"Target Layout BTR: {targetLayoutBtrId}");

                        int fileIndex = 1;
                        foreach (var source in config.SourceFiles)
                        {
                            if (!File.Exists(source.Path)) continue;

                            AcadLogger.LogInfo($"Processing: {Path.GetFileName(source.Path)} -> Layout '{targetLayoutName}'");

                            var sourceDb = new Database(false, true);
                            sourceDb.ReadDwgFile(source.Path, FileShare.ReadWrite, true, "");
                            BindXrefsSafe(sourceDb);

                            using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                            {
                                // Phase 1 Fix: Rename blocks to avoid conflicts
                                RenameBlocksInDb(sourceDb, sourceTrans, $"SL{fileIndex}_");

                                var sourcePsr = GetSourcePaperSpace(sourceDb, sourceTrans);
                                if (sourcePsr == null)
                                {
                                    AcadLogger.LogWarning($"No PaperSpace found in {Path.GetFileName(source.Path)}");
                                    sourceDb.Dispose();
                                    fileIndex++;
                                    continue;
                                }

                                var ids = new ObjectIdCollection();
                                foreach (ObjectId entId in sourcePsr)
                                {
                                    var ent = sourceTrans.GetObject(entId, OpenMode.ForRead) as Entity;
                                    if (ent != null && !string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        ids.Add(entId);
                                    }
                                }

                                if (ids.Count > 0)
                                {
                                    var idMap = new IdMapping();
                                    // Clone to Target Layout BTR instead of ModelSpace
                                    sourceDb.WblockCloneObjects(ids, targetLayoutBtrId, idMap, DuplicateRecordCloning.Replace, false);

                                    var clonedIds = new List<ObjectId>();
                                    foreach (ObjectId id in ids)
                                    {
                                        if (idMap.Contains(id) && !idMap[id].Value.IsNull)
                                            clonedIds.Add(idMap[id].Value);
                                    }
                                    allClonedIds.Add(clonedIds);

                                    var cacheKey = GetExtentsCacheKey(source.Path, source.Layout ?? "Layout1");
                                    allSourceExtents.Add(GetExtents(sourcePsr, cacheKey));
                                    allSourceDbs.Add(sourceDb);
                                }
                                else
                                {
                                    sourceDb.Dispose();
                                }
                                sourceTrans.Commit();
                            }
                            fileIndex++;
                        }

                        // Calculate offsets and transform
                        double maxGlobalHeight = 0;
                        foreach (var ext in allSourceExtents)
                        {
                            double h = ext.MaxPoint.Y - ext.MinPoint.Y;
                            if (h > maxGlobalHeight) maxGlobalHeight = h;
                        }

                        double xOffset = 0;
                        for (int idx = 0; idx < allClonedIds.Count; idx++)
                        {
                            var clonedIds = allClonedIds[idx];
                            var ext = allSourceExtents[idx];
                            double width = ext.MaxPoint.X - ext.MinPoint.X;
                            double height = ext.MaxPoint.Y - ext.MinPoint.Y;

                            double yOffset = 0;
                            if (string.Equals(config.VerticalAlign, "Center", StringComparison.OrdinalIgnoreCase))
                            {
                                yOffset = (maxGlobalHeight - height) / 2.0;
                            }
                            else if (string.Equals(config.VerticalAlign, "Bottom", StringComparison.OrdinalIgnoreCase))
                            {
                                yOffset = maxGlobalHeight - height;
                            }

                            var singleTransform = Matrix3d.Displacement(new Vector3d(xOffset - ext.MinPoint.X, yOffset - ext.MinPoint.Y, 0));
                            foreach (ObjectId destId in clonedIds)
                            {
                                try
                                {
                                    var ent = outputTrans.GetObject(destId, OpenMode.ForWrite) as Entity;
                                    if (ent != null)
                                        ent.TransformBy(singleTransform);
                                }
                                catch { }
                            }

                            xOffset += width + LayoutSpacing;
                            AcadLogger.LogInfo($"Transformed sheet {idx + 1}, offset=({xOffset:F2}, {yOffset:F2})");
                        }

                        outputTrans.Commit();
                    }

                    for (int i = 0; i < allSourceDbs.Count; i++)
                    {
                        try { allSourceDbs[i].Dispose(); } catch { }
                    }

                    CleanupDefaultLayouts(outputDb);
                    RegenerateLayouts(outputDb, "SingleLayout");

                    var dwgVersion = GetDwgVersion(config.DwgVersion);
                    outputDb.SaveAs(config.OutputPath, dwgVersion);
                    ClearExtentsCache();
                    AcadLogger.LogInfo($"SingleLayout completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"SingleLayout error: {ex.Message}");
                AcadLogger.LogError($"Stack: {ex.StackTrace}");
                return false;
            }
        }

        public bool MergeToModelSpace(MergeConfig config)
        {
            try
            {
                AcadLogger.LogSection("MergeToModelSpace");
                AcadLogger.LogInfo($"Output path: {config.OutputPath}");
                AcadLogger.LogInfo($"Source files: {config.SourceFiles?.Count ?? 0}");
                AcadLogger.LogInfo($"Arrangement: {config.ModelSpaceArrangement}");

                if (config.SourceFiles == null || config.SourceFiles.Count == 0)
                {
                    AcadLogger.LogError("No source files provided");
                    return false;
                }

                var validSources = config.SourceFiles.Where(f => f != null && File.Exists(f.Path)).ToList();
                if (validSources.Count == 0)
                {
                    AcadLogger.LogError("No valid source file for ModelSpace merge");
                    return false;
                }

                var outputDb = new Database(true, true);

                using (outputDb)
                {
                    double nextSheetX = 0.0;
                    double nextSheetY = 0.0;
                    double maxRowHeight = 0.0;
                    int currentColumn = 0;
                    int sheetIndex = 0;
                    double spacing = config.CustomSpacing > 0 ? config.CustomSpacing : ModelSpaceSheetMinGap;

                    using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                    {
                        EnsurePaperBackgroundLayer(outputDb, outputTrans);

                        var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(outputDb);
                        var modelSpace = (BlockTableRecord)outputTrans.GetObject(modelSpaceId, OpenMode.ForWrite);

                        foreach (var source in validSources)
                        {
                            sheetIndex++;
                            string label = string.IsNullOrWhiteSpace(source.Layout)
                                ? Path.GetFileNameWithoutExtension(source.Path)
                                : source.Layout;

                            AcadLogger.LogSection($"MODELSPACE SHEET {sheetIndex}/{validSources.Count}: {label}");
                            AcadLogger.LogInfo($"Source: {source.Path}");

                            var sourceDb = new Database(false, true);
                            using (sourceDb)
                            {
                                sourceDb.ReadDwgFile(source.Path, FileShare.ReadWrite, true, "");
                                sourceDb.CloseInput(true);
                                BindXrefsSafe(sourceDb);

                                using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                                {
                                    RenameBlocksInDb(sourceDb, sourceTrans, $"MS{sheetIndex}_");

                                    var sourceLayout = GetSourceLayout(sourceDb, sourceTrans, source.Layout);
                                    var sourcePaperSpace = sourceLayout != null
                                        ? (BlockTableRecord)sourceTrans.GetObject(sourceLayout.BlockTableRecordId, OpenMode.ForRead)
                                        : GetSourcePaperSpace(sourceDb, sourceTrans);

                                    var sourceViewports = CollectModelViewportInfos(
                                        sourceTrans,
                                        sourcePaperSpace,
                                        $"MODELSPACE source viewports: {label}");

                                    int paperExtentEntityCount;
                                    var sourcePaperBounds = GetModelSpaceSheetBounds(
                                        sourceLayout,
                                        sourcePaperSpace,
                                        sourceTrans,
                                        sourceViewports,
                                        out paperExtentEntityCount);

                                    double sheetWidth = Math.Max(1.0, sourcePaperBounds.MaxPoint.X - sourcePaperBounds.MinPoint.X);
                                    double sheetHeight = Math.Max(1.0, sourcePaperBounds.MaxPoint.Y - sourcePaperBounds.MinPoint.Y);

                                    Vector3d placement;
                                    switch (config.ModelSpaceArrangement)
                                    {
                                        case ArrangementMode.Vertical:
                                            placement = new Vector3d(
                                                -sourcePaperBounds.MinPoint.X,
                                                nextSheetY - sourcePaperBounds.MinPoint.Y,
                                                0.0);
                                            break;

                                        case ArrangementMode.Grid:
                                            placement = new Vector3d(
                                                nextSheetX - sourcePaperBounds.MinPoint.X,
                                                nextSheetY - sourcePaperBounds.MinPoint.Y,
                                                0.0);
                                            break;

                                        case ArrangementMode.Horizontal:
                                        default:
                                            placement = new Vector3d(
                                                nextSheetX - sourcePaperBounds.MinPoint.X,
                                                -sourcePaperBounds.MinPoint.Y,
                                                0.0);
                                            break;
                                    }

                                    AcadLogger.LogInfo(
                                        $"MODELSPACE placement: label='{label}', arrangement={config.ModelSpaceArrangement}, " +
                                        $"bounds={FormatExtents(sourcePaperBounds)}, " +
                                        $"paperExtentEntities={paperExtentEntityCount}, placement={FormatVector(placement)}");

                                    var placedIds = new List<ObjectId>();
                                    int backgroundCount = 0;

                                    int paperCloneCount = ClonePaperEntitiesToModelSpace(
                                        sourceDb,
                                        sourceTrans,
                                        outputDb,
                                        outputTrans,
                                        sourcePaperSpace,
                                        modelSpace,
                                        placedIds,
                                        label);

                                    int bakedCount = 0;
                                    if (sourceViewports.Count > 0)
                                    {
                                        bakedCount = BakeModelViewsToPaperSpace(
                                            sourceDb,
                                            sourceTrans,
                                            outputDb,
                                            outputTrans,
                                            modelSpace,
                                            sourceViewports,
                                            label,
                                            placedIds);
                                    }
                                    else
                                    {
                                        AcadLogger.LogWarning($"MODELSPACE: '{label}' has no usable model viewport; only paper-space content was cloned");
                                    }

                                    int movedCount = TransformEntities(outputTrans, placedIds, Matrix3d.Displacement(placement));
                                    MoveModelSpaceBackgroundsToBottom(modelSpace, outputTrans);

                                    AcadLogger.LogInfo(
                                        $"MODELSPACE summary: '{label}' background={backgroundCount}, paperClones={paperCloneCount}, " +
                                        $"bakedModelClones={bakedCount}, moved={movedCount}, " +
                                        $"position=({nextSheetX:F2},{nextSheetY:F2}), size=({sheetWidth:F2},{sheetHeight:F2})");

                                    double gap = Math.Max(spacing, sheetWidth * 0.02);
                                    double vGap = Math.Max(spacing, sheetHeight * 0.02);

                                    switch (config.ModelSpaceArrangement)
                                    {
                                        case ArrangementMode.Vertical:
                                            nextSheetY += sheetHeight + vGap;
                                            break;

                                        case ArrangementMode.Grid:
                                            nextSheetX += sheetWidth + gap;
                                            maxRowHeight = Math.Max(maxRowHeight, sheetHeight);
                                            currentColumn++;

                                            if (currentColumn >= config.GridColumns)
                                            {
                                                currentColumn = 0;
                                                nextSheetX = 0;
                                                nextSheetY += maxRowHeight + vGap;
                                                maxRowHeight = 0;
                                            }
                                            break;

                                        case ArrangementMode.Horizontal:
                                        default:
                                            nextSheetX += sheetWidth + gap;
                                            break;
                                    }

                                    sourceTrans.Commit();
                                }
                            }
                        }

                        outputTrans.Commit();
                    }

                    CleanupDefaultLayouts(outputDb);
                    RegenerateModelSpace(outputDb);

                    var dwgVersion = GetDwgVersion(config.DwgVersion);
                    outputDb.SaveAs(config.OutputPath, dwgVersion);
                    ClearExtentsCache();
                    AcadLogger.Log($"[LayoutMerger] ModelSpace merge completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] ModelSpace error: {ex.Message}");
                AcadLogger.Log($"[LayoutMerger] ModelSpace stack: {ex.StackTrace}");
                return false;
            }
        }

        public bool VerifyCombinedFile(MergeConfig config, out string message)
        {
            message = null;

            try
            {
                if (config == null)
                {
                    message = "Verification failed: merge config is null.";
                    AcadLogger.LogError(message);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(config.OutputPath) || !File.Exists(config.OutputPath))
                {
                    message = $"Verification failed: output DWG not found: {config.OutputPath}";
                    AcadLogger.LogError(message);
                    return false;
                }

                var fileInfo = new FileInfo(config.OutputPath);
                if (fileInfo.Length < 4096)
                {
                    message = $"Verification failed: output DWG is too small ({fileInfo.Length} bytes).";
                    AcadLogger.LogError(message);
                    return false;
                }

                int expected = config.ExpectedSheetCount > 0
                    ? config.ExpectedSheetCount
                    : (config.SourceFiles?.Count ?? 0);

                var db = new Database(false, true);
                using (db)
                {
                    db.ReadDwgFile(config.OutputPath, FileShare.ReadWrite, true, "");
                    db.CloseInput(true);

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        if (string.Equals(config.Mode, "MultiLayout", StringComparison.OrdinalIgnoreCase))
                        {
                            var layoutStats = InspectPaperLayouts(db, tr);
                            var contentLayouts = layoutStats
                                .Where(s => s.ContentEntityCount > 0)
                                .ToList();
                            var emptyNonDefaultLayouts = layoutStats
                                .Where(s => s.ContentEntityCount == 0 && !IsDefaultLayoutName(s.Name))
                                .Select(s => s.Name)
                                .ToList();
                            var emptyDefaultLayouts = layoutStats
                                .Where(s => s.ContentEntityCount == 0 && IsDefaultLayoutName(s.Name))
                                .Select(s => s.Name)
                                .ToList();

                            AcadLogger.LogInfo(
                                $"VERIFY MultiLayout: layouts={layoutStats.Count}, contentLayouts={contentLayouts.Count}, " +
                                $"emptyDefaultLayouts={emptyDefaultLayouts.Count}, expected={expected}");

                            foreach (var stat in layoutStats)
                            {
                                AcadLogger.LogInfo(
                                    $"VERIFY layout '{stat.Name}': entities={stat.EntityCount}, contentEntities={stat.ContentEntityCount}, " +
                                    $"backgroundEntities={stat.BackgroundEntityCount}, viewportEntities={stat.ViewportEntityCount}");
                            }

                            if (emptyDefaultLayouts.Count > 0)
                                AcadLogger.LogWarning("VERIFY ignored empty default layout(s): " + string.Join(", ", emptyDefaultLayouts));

                            if (expected > 0 && contentLayouts.Count < expected)
                            {
                                message = $"Verification failed: expected {expected} content layout(s), found {contentLayouts.Count}.";
                                AcadLogger.LogError(message);
                                return false;
                            }

                            if (emptyNonDefaultLayouts.Count > 0)
                            {
                                message = "Verification failed: empty non-default layout(s): " + string.Join(", ", emptyNonDefaultLayouts);
                                AcadLogger.LogError(message);
                                return false;
                            }
                        }
                        else
                        {
                            var modelStats = GetModelSpaceStats(db, tr);
                            LogModelSpaceStats($"VERIFY {config.Mode}", modelStats);

                            if (modelStats.EntityCount == 0 || modelStats.ExtentsEntityCount == 0)
                            {
                                message = $"Verification failed: {config.Mode} ModelSpace has no usable entities.";
                                AcadLogger.LogError(message);
                                return false;
                            }

                            if (string.Equals(config.Mode, "ModelSpace", StringComparison.OrdinalIgnoreCase))
                            {
                                int backgroundCount = CountModelSpaceBackgrounds(db, tr);
                                AcadLogger.LogInfo($"VERIFY ModelSpace: sheetBackgrounds={backgroundCount}, expected={expected}");

                                if (expected > 0 && backgroundCount < expected)
                                {
                                    message = $"Verification failed: expected {expected} model-space sheet region(s), found {backgroundCount}.";
                                    AcadLogger.LogError(message);
                                    return false;
                                }
                            }
                        }

                        tr.Commit();
                    }
                }

                message = "Verification passed.";
                AcadLogger.LogInfo(message);
                return true;
            }
            catch (System.Exception ex)
            {
                message = "Verification failed: " + ex.Message;
                AcadLogger.LogError(message);
                AcadLogger.LogError("Verification stack: " + ex.StackTrace);
                return false;
            }
        }

        public void CreateCombinedDwgIndex(MergeConfig config)
        {
            try
            {
                if (config == null || !config.SheetSetEnabled)
                    return;

                if (string.IsNullOrWhiteSpace(config.SheetSetIndexPath))
                    return;

                var index = new
                {
                    Type = "Licorp Combined DWG Index",
                    Note = "This is a lightweight JSON index for the combined DWG layouts. It is not an AutoCAD Sheet Set Manager DST file.",
                    CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    OutputDwg = config.OutputPath,
                    Mode = config.Mode,
                    SheetCount = config.SourceFiles?.Count ?? 0,
                    Sheets = (config.SourceFiles ?? new List<SourceFile>())
                        .Select((s, i) => new
                        {
                            Index = i + 1,
                            Layout = s.Layout,
                            SourceDwg = s.Path,
                            Region = string.Equals(config.Mode, "ModelSpace", StringComparison.OrdinalIgnoreCase)
                                ? $"ModelSpace sheet region {i + 1}"
                                : s.Layout
                        })
                        .ToList()
                };

                var folder = Path.GetDirectoryName(config.SheetSetIndexPath);
                if (!string.IsNullOrWhiteSpace(folder))
                    Directory.CreateDirectory(folder);

                File.WriteAllText(config.SheetSetIndexPath, JsonConvert.SerializeObject(index, Formatting.Indented));
                AcadLogger.LogInfo($"Combined DWG index written: {config.SheetSetIndexPath}");

                // Create actual DST file for AutoCAD Sheet Set Manager
                var dstPath = Path.ChangeExtension(config.SheetSetIndexPath, ".dst");
                CreateSheetSetDstFile(dstPath, config);
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Combined DWG index failed: {ex.Message}");
            }
        }

        private void CreateSheetSetDstFile(string dstPath, MergeConfig config)
        {
            try
            {
                var sheetSetName = Path.GetFileNameWithoutExtension(config.OutputPath ?? "Combined");
                var sheets = config.SourceFiles ?? new List<SourceFile>();
                var dwgRelativePath = Path.GetFileName(config.OutputPath ?? "output.dwg");

                var sb = new System.Text.StringBuilder();
                sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sb.AppendLine("<!-- AutoCAD Sheet Set Manager DST file -->");
                sb.AppendLine("<SheetSet xmlns=\"http://www.autodesk.com/SheetSet\">");
                sb.AppendLine($"  <Name>{EscapeXml(sheetSetName)}</Name>");
                sb.AppendLine($"  <Description>Merged from {sheets.Count} sheets by Licorp CombineCAD</Description>");
                sb.AppendLine($"  <OutputDwg>{EscapeXml(dwgRelativePath)}</OutputDwg>");
                sb.AppendLine($"  <CreatedAt>{DateTime.Now:yyyy-MM-dd HH:mm:ss}</CreatedAt>");
                sb.AppendLine("  <Sheets>");

                for (int i = 0; i < sheets.Count; i++)
                {
                    var s = sheets[i];
                    var layoutName = s.Layout ?? $"Layout{i + 1}";
                    var sourceName = Path.GetFileNameWithoutExtension(s.Path ?? "");
                    sb.AppendLine($"    <Sheet>");
                    sb.AppendLine($"      <Index>{i + 1}</Index>");
                    sb.AppendLine($"      <Name>{EscapeXml(layoutName)}</Name>");
                    sb.AppendLine($"      <Description>From: {EscapeXml(sourceName)}</Description>");
                    sb.AppendLine($"      <Layout>{EscapeXml(layoutName)}</Layout>");
                    sb.AppendLine($"      <DwgFile>{EscapeXml(dwgRelativePath)}</DwgFile>");
                    sb.AppendLine($"    </Sheet>");
                }

                sb.AppendLine("  </Sheets>");
                sb.AppendLine("</SheetSet>");

                File.WriteAllText(dstPath, sb.ToString(), System.Text.Encoding.UTF8);
                AcadLogger.LogInfo($"Sheet Set DST file created: {dstPath}");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Sheet Set DST creation failed: {ex.Message}");
            }
        }

        private static string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }

        public void HandleRasterImages(MergeConfig config)
        {
            try
            {
                if (config == null || string.IsNullOrWhiteSpace(config.OutputPath) || !File.Exists(config.OutputPath))
                    return;

                var rasterInfos = ScanRasterImages(config.OutputPath);
                AcadLogger.LogInfo($"Raster image scan: mode={config.RasterImageMode}, count={rasterInfos.Count}");

                if (rasterInfos.Count == 0)
                {
                    AcadLogger.LogInfo("No raster image entities found.");
                    return;
                }

                // Mode: CopyAlongside - copy raster files to output folder
                if (string.Equals(config.RasterImageMode, "CopyAlongside", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(config.RasterImageMode, "KeepReference", StringComparison.OrdinalIgnoreCase))
                {
                    CopyRasterFilesAlongside(config, rasterInfos);
                    return;
                }

                // Mode: EmbedAsOle - attempt OLE embed, fallback to copy
                if (string.Equals(config.RasterImageMode, "EmbedAsOle", StringComparison.OrdinalIgnoreCase))
                {
                    AcadLogger.LogInfo("OLE embed not available in .NET API; falling back to copy alongside.");
                    CopyRasterFilesAlongside(config, rasterInfos);
                    return;
                }

                AcadLogger.LogInfo($"Raster image handling: mode={config.RasterImageMode}, no action taken.");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Raster image handling failed: {ex.Message}");
            }
        }

        private void CopyRasterFilesAlongside(MergeConfig config, List<RasterImageInfo> rasterInfos)
        {
            var outputDir = Path.GetDirectoryName(config.OutputPath);
            var sourceFileDirs = (config.SourceFiles ?? new List<SourceFile>())
                .Where(s => !string.IsNullOrEmpty(s.Path))
                .Select(s => Path.GetDirectoryName(s.Path))
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            int copiedCount = 0;
            var copiedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var rasterExtensions = new[] { ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif", ".pcx", ".tga", ".rlc" };

            foreach (var srcDir in sourceFileDirs)
            {
                try
                {
                    foreach (var ext in rasterExtensions)
                    {
                        var rasterFiles = Directory.GetFiles(srcDir, "*" + ext, SearchOption.TopDirectoryOnly);
                        foreach (var rasterFile in rasterFiles)
                        {
                            var fileName = Path.GetFileName(rasterFile);
                            if (copiedFiles.Contains(fileName))
                                continue;

                            var destPath = Path.Combine(outputDir, fileName);
                            if (!File.Exists(destPath))
                            {
                                File.Copy(rasterFile, destPath, false);
                                copiedCount++;
                            }
                            copiedFiles.Add(fileName);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"Raster copy from '{srcDir}' failed: {ex.Message}");
                }
            }

            AcadLogger.LogInfo($"Raster files copied alongside DWG: {copiedCount} file(s) to {outputDir}");
        }

        // ============ HELPER METHODS ============

        private void ApplyPaperBackgroundPresentation(Database db, string mode)
        {
            try
            {
                AcadLogger.LogSection($"Paper Background ({mode})");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    EnsurePaperBackgroundLayer(db, tr);
                    int backgroundCount = AddWhitePaperBackgrounds(db, tr);

                    tr.Commit();

                    AcadLogger.LogInfo($"PAPER BACKGROUND: whiteBackgrounds={backgroundCount}");
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Paper background failed for {mode}: {ex.Message}");
            }
        }

        private void RemovePaperBackgroundPresentation(Database db, string mode)
        {
            try
            {
                AcadLogger.LogSection($"Paper Background Cleanup ({mode})");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    int erased = ErasePaperBackgrounds(db, tr);
                    tr.Commit();

                    AcadLogger.LogInfo($"PAPER BACKGROUND: erased={erased}, mode={mode}");
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Paper background cleanup failed for {mode}: {ex.Message}");
            }
        }

        private ObjectId EnsurePaperBackgroundLayer(Database db, Transaction tr)
        {
            var white = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
            var layers = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (!layers.Has(PaperBackgroundLayerName))
            {
                layers.UpgradeOpen();
                var layer = new LayerTableRecord
                {
                    Name = PaperBackgroundLayerName,
                    Color = white
                };

                ObjectId layerId = layers.Add(layer);
                tr.AddNewlyCreatedDBObject(layer, true);
                AcadLogger.LogInfo($"PAPER BACKGROUND: created layer '{PaperBackgroundLayerName}'");
                return layerId;
            }

            ObjectId existingId = layers[PaperBackgroundLayerName];
            var existingLayer = (LayerTableRecord)tr.GetObject(existingId, OpenMode.ForWrite);
            existingLayer.Color = white;
            return existingId;
        }

        private ObjectId EnsureViewportLayer(Database db, Transaction tr)
        {
            var layers = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 8);

            if (layers.Has(ViewportLayerName))
            {
                var id = layers[ViewportLayerName];
                var layer = (LayerTableRecord)tr.GetObject(id, OpenMode.ForWrite);
                layer.IsPlottable = false;
                layer.Color = color;
                return id;
            }

            layers.UpgradeOpen();
            var newLayer = new LayerTableRecord
            {
                Name = ViewportLayerName,
                IsPlottable = false,
                Color = color
            };
            var layerId = layers.Add(newLayer);
            tr.AddNewlyCreatedDBObject(newLayer, true);
            AcadLogger.LogInfo($"VIEWPORT LAYER: created non-plot layer '{ViewportLayerName}'");
            return layerId;
        }

        private int MergeLayersFromSource(Database sourceDb, Database outputDb, Transaction sourceTr, Transaction outputTr)
        {
            int mergedCount = 0;
            try
            {
                var srcLayers = (LayerTable)sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead);
                var outLayers = (LayerTable)outputTr.GetObject(outputDb.LayerTableId, OpenMode.ForRead);

                foreach (ObjectId srcLayerId in srcLayers)
                {
                    try
                    {
                        var srcLayer = (LayerTableRecord)sourceTr.GetObject(srcLayerId, OpenMode.ForRead);
                        string layerName = srcLayer.Name;

                        if (string.IsNullOrEmpty(layerName))
                            continue;

                        if (!outLayers.Has(layerName))
                        {
                            outLayers.UpgradeOpen();
                            var newLayer = new LayerTableRecord
                            {
                                Name = layerName,
                                Color = srcLayer.Color,
                                IsPlottable = srcLayer.IsPlottable,
                                IsFrozen = false,
                                IsOff = false
                            };

                            try { newLayer.LineWeight = srcLayer.LineWeight; } catch { }
                            try { newLayer.LinetypeObjectId = srcLayer.LinetypeObjectId; } catch { }

                            outLayers.Add(newLayer);
                            outputTr.AddNewlyCreatedDBObject(newLayer, true);
                            mergedCount++;
                        }
                    }
                    catch { }
                }

                if (mergedCount > 0)
                    AcadLogger.LogInfo($"Layer merge: {mergedCount} new layer(s) added from source");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Layer merge partial failure: {ex.Message}");
            }
            return mergedCount;
        }

        private int AddWhitePaperBackgrounds(Database db, Transaction tr)
        {
            int added = 0;
            var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layouts)
            {
                if (string.Equals(entry.Key, "Model", StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (layout.BlockTableRecordId.IsNull)
                        continue;

                    var paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                    int erasedOld = EraseExistingPaperBackgrounds(paperSpace, tr);
                    int extentEntities;
                    var backgroundExtents = GetPaperBackgroundExtents(layout, paperSpace, tr, out extentEntities);
                    var backgroundIds = CreateWhiteBackgroundHatch(db, tr, paperSpace, backgroundExtents);
                    bool drawOrderMoved = MoveEntitiesToBottom(paperSpace, tr, backgroundIds);
                    EnsureLayoutViewportsOnTop(tr, paperSpace, entry.Key);

                    added++;
                    AcadLogger.LogInfo(
                        $"PAPER BACKGROUND: layout '{entry.Key}' white background added, " +
                        $"erasedOld={erasedOld}, extentsEntities={extentEntities}, " +
                        $"background={FormatExtents(backgroundExtents)}, drawOrderBottom={drawOrderMoved}");
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"PAPER BACKGROUND: white background failed for layout '{entry.Key}': {ex.Message}");
                }
            }

            return added;
        }

        private int EraseExistingPaperBackgrounds(BlockTableRecord paperSpace, Transaction tr)
        {
            var idsToErase = new List<ObjectId>();

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent != null &&
                        !ent.IsErased &&
                        string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        idsToErase.Add(id);
                    }
                }
                catch
                {
                }
            }

            int erased = 0;
            foreach (ObjectId id in idsToErase)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                    if (ent != null && !ent.IsErased)
                    {
                        ent.Erase();
                        erased++;
                    }
                }
                catch
                {
                }
            }

            return erased;
        }

        private int ErasePaperBackgrounds(Database db, Transaction tr)
        {
            int erased = 0;
            var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layouts)
            {
                if (string.Equals(entry.Key, "Model", StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (layout == null || layout.BlockTableRecordId.IsNull)
                        continue;

                    var paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                    erased += EraseExistingPaperBackgrounds(paperSpace, tr);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"PAPER BACKGROUND: cleanup failed for layout '{entry.Key}': {ex.Message}");
                }
            }

            return erased;
        }

        private Extents3d GetLayoutPaperFallbackExtents(Layout layout, BlockTableRecord paperSpace, Transaction tr)
        {
            int extentEntities;
            string cacheKey = layout != null ? GetExtentsCacheKey(null, layout.LayoutName + "_fallback") : null;
            var bounds = GetPaperEntityExtentsExcludingViewports(paperSpace, tr, out extentEntities, cacheKey);
            bool hasBounds = IsUsableExtents(bounds);

            double maxViewArea = 0.0;
            var rawVps = new List<Viewport>();

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var vp = tr.GetObject(id, OpenMode.ForRead, false) as Viewport;
                    if (!IsRawModelViewport(vp))
                        continue;

                    rawVps.Add(vp);
                    maxViewArea = Math.Max(maxViewArea, GetViewportViewArea(vp));
                }
                catch
                {
                }
            }

            foreach (var vp in rawVps)
            {
                try
                {
                    if (IsUtilityViewport(vp, maxViewArea))
                        continue;

                    var vpInfo = new ViewportInfo
                    {
                        CenterPoint = vp.CenterPoint,
                        Width = vp.Width,
                        Height = vp.Height
                    };

                    var vpPaper = GetViewportPaperExtents(vpInfo);
                    bounds = hasBounds ? CombineExtents(bounds, vpPaper) : vpPaper;
                    hasBounds = true;
                }
                catch
                {
                }
            }

            if (layout != null && layout.PlotPaperSize.X > 1.0 && layout.PlotPaperSize.Y > 1.0)
            {
                var plotBounds = new Extents3d(
                    new Point3d(0.0, 0.0, 0.0),
                    new Point3d(layout.PlotPaperSize.X, layout.PlotPaperSize.Y, 0.0));
                bounds = hasBounds ? CombineExtents(bounds, plotBounds) : plotBounds;
                hasBounds = true;
            }

            if (!hasBounds)
            {
                bounds = new Extents3d(
                    new Point3d(0.0, 0.0, 0.0),
                    new Point3d(PaperBackgroundFallbackWidth, PaperBackgroundFallbackHeight, 0.0));
            }

            return bounds;
        }

        private PaperContextResult EnsureLayoutPaperContextFromGeometry(
            Database db,
            Transaction tr,
            Layout layout,
            BlockTableRecord paperSpace,
            string layoutName,
            string phase,
            string revitPaperSize = null)
        {
            var result = new PaperContextResult
            {
                AppliedRotation = PlotRotation.Degrees000,
                RequiredWidth = 0.0,
                RequiredHeight = 0.0,
                WasAdjusted = false
            };

            try
            {
                if (layout == null || paperSpace == null)
                    return result;

                int extentEntities;
                string paperCacheKey = GetExtentsCacheKey(null, layoutName + "_paper");
                var contentBounds = GetPaperEntityExtentsExcludingViewports(paperSpace, tr, out extentEntities, paperCacheKey);
                bool hasContentBounds = IsUsableExtents(contentBounds);

                // Normalize title-sheet geometry to paper origin before resolving page setup.
                // This avoids inflated paper sizes/origins when imported sheets are offset.
                if (hasContentBounds)
                {
                    var shiftX = -contentBounds.MinPoint.X;
                    var shiftY = -contentBounds.MinPoint.Y;

                    if (Math.Abs(shiftX) > 1e-6 || Math.Abs(shiftY) > 1e-6)
                    {
                        var moved = TranslatePaperEntitiesToOrigin(paperSpace, tr, shiftX, shiftY);
                        result.ShiftX = shiftX;
                        result.ShiftY = shiftY;
                        AcadLogger.LogInfo(
                            $"{phase}: normalized paper entities to origin for '{layoutName}', " +
                            $"move=({shiftX:F4},{shiftY:F4}), moved={moved}");

                        // Recompute after transform so the selected media matches final geometry.
                        contentBounds = GetPaperEntityExtentsExcludingViewports(paperSpace, tr, out extentEntities, paperCacheKey + "_after");
                        hasContentBounds = IsUsableExtents(contentBounds);
                    }
                }

                var fallback = hasContentBounds
                    ? contentBounds
                    : GetLayoutPaperFallbackExtents(layout, paperSpace, tr);

                if (hasContentBounds)
                {
                    var normalizedBounds = NormalizeLayoutPaperToTitleSheet(
                        db, layout, contentBounds, layoutName, phase, extentEntities, revitPaperSize);
                    if (IsUsableExtents(normalizedBounds))
                        fallback = normalizedBounds;

                    AcadLogger.LogInfo(
                        $"{phase}: using title sheet extents for paper context '{layoutName}' => " +
                        $"bounds={FormatExtents(fallback)}, extentEntities={extentEntities}");
                }

                double width = Math.Max(1.0, fallback.MaxPoint.X - fallback.MinPoint.X);
                double height = Math.Max(1.0, fallback.MaxPoint.Y - fallback.MinPoint.Y);

                double currentWidth = layout.PlotPaperSize.X;
                double currentHeight = layout.PlotPaperSize.Y;

                double tol = 2.0; // mm
                bool invalidPlotSize = currentWidth < 1.0 || currentHeight < 1.0;

                // Paper may have been rotated 90° by NormalizeLayoutPaperToTitleSheet,
                // so check both orientations: paper (W,H) or (H,W) must cover title sheet.
                bool fitsNormal = currentWidth >= width - tol && currentHeight >= height - tol;
                bool fitsRotated = currentWidth >= height - tol && currentHeight >= width - tol;
                bool mismatch = !fitsNormal && !fitsRotated;

                result.RequiredWidth = width;
                result.RequiredHeight = height;

                if (!invalidPlotSize && !mismatch)
                {
                    AcadLogger.LogInfo(
                        $"{phase}: paper context exact for {layoutName} " +
                        $"paper={currentWidth:F2},{currentHeight:F2}, title={width:F2},{height:F2}");
                    try { result.AppliedRotation = layout.PlotRotation; } catch { }
                    result.WasAdjusted = result.AppliedRotation != PlotRotation.Degrees000;
                    return result;
                }

                AcadLogger.LogInfo(
                    $"{phase}: Page setup paper mismatch (auto-adjusting) for '{layoutName}', " +
                    $"paper=({currentWidth:F2}, {currentHeight:F2}), title=({width:F2}, {height:F2}), " +
                    $"bounds={FormatExtents(fallback)}, extentEntities={extentEntities}");

                try
                {
                    EnsurePlotSettingsRefreshed(layout);
                    var psv = PlotSettingsValidator.Current;

                    // PlotType.Layout + media selection
                    string selectedMedia = SelectBestCanonicalMedia(psv, layout, width, height);
                    if (!string.IsNullOrWhiteSpace(selectedMedia))
                    {
                        psv.SetCanonicalMediaName(layout, selectedMedia);
                        AcadLogger.LogInfo(
                            $"{phase}: Page setup media adjusted for '{layoutName}' => " +
                            $"media='{selectedMedia}', paper=({layout.PlotPaperSize.X:F2}, {layout.PlotPaperSize.Y:F2})");
                    }

                    // FIX #4: Validate paper size after selection and try to correct if mismatch
                    double actualPaperW = layout.PlotPaperSize.X;
                    double actualPaperH = layout.PlotPaperSize.Y;
                    double sizeTolerance = 50.0;
                    bool sizeMismatch = Math.Abs(actualPaperW - width) > sizeTolerance ||
                                        Math.Abs(actualPaperH - height) > sizeTolerance;

                    if (sizeMismatch && !string.IsNullOrWhiteSpace(selectedMedia))
                    {
                        AcadLogger.LogWarning(
                            $"{phase}: PAPER SIZE MISMATCH for '{layoutName}': " +
                            $"required=({width:F0}x{height:F0}), " +
                            $"actual=({actualPaperW:F0}x{actualPaperH:F0}), " +
                            $"media='{selectedMedia}'. Attempting correction...");

                        TryApplyCorrectedPaperSize(psv, layout, width, height, layoutName, phase);
                    }

                    try { psv.SetPlotType(layout, PlotType.Layout); } catch { }
                    try { psv.SetPlotCentered(layout, false); } catch { }
                    try { psv.SetPlotOrigin(layout, new Point2d(0.0, 0.0)); } catch { }

                    try { result.AppliedRotation = layout.PlotRotation; } catch { }
                    result.WasAdjusted = true;

                    AcadLogger.LogInfo(
                        $"{phase}: geometry paper context applied for '{layoutName}', " +
                        $"required=({width:F2}, {height:F2}), rotation={result.AppliedRotation}");
                }
                catch (System.Exception psvEx)
                {
                    AcadLogger.LogWarning(
                        $"{phase}: PlotSettingsValidator fallback failed for '{layoutName}': {psvEx.Message}");
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning(
                    $"{phase}: EnsureLayoutPaperContextFromGeometry failed for '{layoutName}': {ex.Message}");
            }

            return result;
        }

        private int TranslatePaperEntitiesToOrigin(BlockTableRecord paperSpace, Transaction tr, double dx, double dy)
        {
            int moved = 0;
            var transform = Matrix3d.Displacement(new Vector3d(dx, dy, 0.0));

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                    if (ent == null || ent.IsErased || ent is Viewport)
                        continue;

                    if (string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    ent.TransformBy(transform);
                    moved++;
                }
                catch
                {
                }
            }

            return moved;
        }

        private Extents3d NormalizeLayoutPaperToTitleSheet(
            Database db,
            Layout layout,
            Extents3d titleBounds,
            string layoutName,
            string phase,
            int extentEntities,
            string revitPaperSize = null)
        {
            try
            {
                if (layout == null || !IsUsableExtents(titleBounds))
                    return titleBounds;

                double minX = titleBounds.MinPoint.X;
                double minY = titleBounds.MinPoint.Y;
                double width = Math.Max(1.0, titleBounds.MaxPoint.X - titleBounds.MinPoint.X);
                double height = Math.Max(1.0, titleBounds.MaxPoint.Y - titleBounds.MinPoint.Y);

                var previousWorkingDb = HostApplicationServices.WorkingDatabase;
                try
                {
                    HostApplicationServices.WorkingDatabase = db;
                    EnsurePlotSettingsRefreshed(layout);
                    var psv = PlotSettingsValidator.Current;

                    // Try to use Revit paper size first for accurate matching
                    string selectedMedia = null;
                    if (!string.IsNullOrWhiteSpace(revitPaperSize))
                    {
                        selectedMedia = FindCanonicalMediaByRevitSize(psv, layout, revitPaperSize, width, height);
                        if (!string.IsNullOrWhiteSpace(selectedMedia))
                            AcadLogger.LogInfo($"{phase}: Using Revit paper size '{revitPaperSize}' -> media='{selectedMedia}' for '{layoutName}'");
                    }

                    // Fallback to geometry-based matching
                    if (string.IsNullOrWhiteSpace(selectedMedia))
                        selectedMedia = SelectBestCanonicalMedia(psv, layout, width, height);

                    if (!string.IsNullOrWhiteSpace(selectedMedia))
                    {
                        psv.SetCanonicalMediaName(layout, selectedMedia);
                        double paperW = layout.PlotPaperSize.X;
                        double paperH = layout.PlotPaperSize.Y;

                        double dNormal = Math.Abs(paperW - width) + Math.Abs(paperH - height);
                        double dRotated = Math.Abs(paperW - height) + Math.Abs(paperH - width);
                        bool rotate = dRotated < dNormal;

                        try { psv.SetPlotRotation(layout, rotate ? PlotRotation.Degrees090 : PlotRotation.Degrees000); } catch { }
                        try { psv.SetPlotType(layout, PlotType.Layout); } catch { }
                        try { psv.SetPlotCentered(layout, false); } catch { }
                        try { psv.SetPlotOrigin(layout, new Point2d(-minX, -minY)); } catch { }

                        AcadLogger.LogInfo(
                            $"{phase}: title sheet paper matched for '{layoutName}', " +
                            $"titleBounds={FormatExtents(titleBounds)}, titleSize=({width:F2},{height:F2}), " +
                            $"media='{selectedMedia}', paper=({layout.PlotPaperSize.X:F2},{layout.PlotPaperSize.Y:F2}), " +
                            $"origin=({-minX:F2},{-minY:F2}), extentEntities={extentEntities}");

                        return new Extents3d(
                            new Point3d(0.0, 0.0, 0.0),
                            new Point3d(width, height, 0.0));
                    }
                }
                finally
                {
                    try { HostApplicationServices.WorkingDatabase = previousWorkingDb; } catch { }
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"{phase}: title sheet paper match failed for '{layoutName}': {ex.Message}");
            }

            return titleBounds;
        }

        // Cache for media sizes to avoid repeated SetCanonicalMediaName calls
        private static Dictionary<string, Tuple<double, double>> _mediaSizeCache = new Dictionary<string, Tuple<double, double>>(StringComparer.OrdinalIgnoreCase);

        // Cache for corrected paper size results: key = "WxH", value = (mediaName, isRotated)
        private static Dictionary<string, Tuple<string, bool>> _correctedPaperCache = new Dictionary<string, Tuple<string, bool>>(StringComparer.OrdinalIgnoreCase);

        private string SelectBestCanonicalMedia(
            PlotSettingsValidator psv,
            Layout layout,
            double requiredWidth,
            double requiredHeight)
        {
            string originalMedia = null;
            try { originalMedia = layout.CanonicalMediaName; } catch { }

            string best = null;
            bool bestIsRotated = false;
            double bestScore = double.MaxValue;
            const double exactTol = 2.0;

            try
            {
                var mediaNames = psv.GetCanonicalMediaNameList(layout);

                if (_mediaSizeCache.Count == 0)
                {
                    foreach (string mediaName in mediaNames)
                    {
                        if (string.IsNullOrWhiteSpace(mediaName) || _mediaSizeCache.ContainsKey(mediaName))
                            continue;

                        try
                        {
                            psv.SetCanonicalMediaName(layout, mediaName);
                            double mw = layout.PlotPaperSize.X;
                            double mh = layout.PlotPaperSize.Y;
                            if (mw >= 1.0 && mh >= 1.0)
                                _mediaSizeCache[mediaName] = Tuple.Create(mw, mh);
                        }
                        catch { }
                    }
                    AcadLogger.LogInfo($"Media size cache populated: {_mediaSizeCache.Count} entries");
                }

                foreach (var kvp in _mediaSizeCache)
                {
                    double mw = kvp.Value.Item1;
                    double mh = kvp.Value.Item2;

                    double dNormal = Math.Abs(mw - requiredWidth) + Math.Abs(mh - requiredHeight);
                    double dRotated = Math.Abs(mw - requiredHeight) + Math.Abs(mh - requiredWidth);

                    if (dNormal < dRotated)
                    {
                        if (dNormal < bestScore)
                        {
                            bestScore = dNormal;
                            best = kvp.Key;
                            bestIsRotated = false;
                        }
                    }
                    else
                    {
                        if (dRotated < bestScore)
                        {
                            bestScore = dRotated;
                            best = kvp.Key;
                            bestIsRotated = true;
                        }
                    }

                    if (bestScore <= exactTol)
                        break;
                }

                if (best != null && _mediaSizeCache.ContainsKey(best))
                {
                    var sz = _mediaSizeCache[best];
                    double pw = sz.Item1;
                    double ph = sz.Item2;
                    if (pw < requiredWidth - 2.0 || ph < requiredHeight - 2.0)
                    {
                        string larger = FindLargerMedia(requiredWidth, requiredHeight);
                        if (!string.IsNullOrWhiteSpace(larger))
                        {
                            AcadLogger.LogInfo(
                                $"MEDIA: Upgraded '{best}' ({pw:F0}x{ph:F0}) to '{larger}' " +
                                $"for required {requiredWidth:F0}x{requiredHeight:F0}");
                            best = larger;
                            bestIsRotated = false;
                        }
                    }
                }

                if (bestIsRotated && best != null)
                {
                    try
                    {
                        psv.SetCanonicalMediaName(layout, best);
                        psv.SetPlotRotation(layout, PlotRotation.Degrees090);
                        AcadLogger.LogInfo(
                            $"MEDIA: Selected '{best}' with 90° rotation for {requiredWidth:F0}x{requiredHeight:F0}, " +
                            $"paper=({layout.PlotPaperSize.X:F0}x{layout.PlotPaperSize.Y:F0})");
                    }
                    catch (System.Exception rotEx)
                    {
                        AcadLogger.LogWarning($"MEDIA: Failed to apply rotation for '{best}': {rotEx.Message}");
                    }
                }
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(originalMedia))
                {
                    try { psv.SetCanonicalMediaName(layout, originalMedia); } catch { }
                }
            }

            return best;
        }

        /// <summary>
        /// Find canonical media name using Revit paper size string (e.g., "A1", "ArchE1", "A0 Landscape").
        /// Maps Revit paper size names to AutoCAD canonical media names by matching against the media cache.
        /// </summary>
        private string FindCanonicalMediaByRevitSize(
            PlotSettingsValidator psv,
            Layout layout,
            string revitPaperSize,
            double fallbackWidth,
            double fallbackHeight)
        {
            if (string.IsNullOrWhiteSpace(revitPaperSize))
                return null;

            string originalMedia = null;
            try { originalMedia = layout.CanonicalMediaName; } catch { }

            try
            {
                // Populate media cache if needed
                if (_mediaSizeCache.Count == 0)
                {
                    var mediaNames = psv.GetCanonicalMediaNameList(layout);
                    foreach (string mediaName in mediaNames)
                    {
                        if (string.IsNullOrWhiteSpace(mediaName) || _mediaSizeCache.ContainsKey(mediaName))
                            continue;
                        try
                        {
                            psv.SetCanonicalMediaName(layout, mediaName);
                            double mw = layout.PlotPaperSize.X;
                            double mh = layout.PlotPaperSize.Y;
                            if (mw >= 1.0 && mh >= 1.0)
                                _mediaSizeCache[mediaName] = Tuple.Create(mw, mh);
                        }
                        catch { }
                    }
                }

                // Normalize Revit paper size for matching
                string normalizedRevit = revitPaperSize.Trim().ToUpperInvariant();

                // Standard paper size dimensions in mm (width, height) for landscape orientation
                var standardSizes = new Dictionary<string, Tuple<double, double>>(StringComparer.OrdinalIgnoreCase)
                {
                    { "A0", Tuple.Create(1189.0, 841.0) },
                    { "A1", Tuple.Create(841.0, 594.0) },
                    { "A2", Tuple.Create(594.0, 420.0) },
                    { "A3", Tuple.Create(420.0, 297.0) },
                    { "A4", Tuple.Create(297.0, 210.0) },
                    { "ANSI A", Tuple.Create(279.4, 215.9) },
                    { "ANSI B", Tuple.Create(431.8, 279.4) },
                    { "ANSI C", Tuple.Create(558.8, 431.8) },
                    { "ANSI D", Tuple.Create(863.6, 558.8) },
                    { "ANSI E", Tuple.Create(1117.6, 863.6) },
                    { "ARCH A", Tuple.Create(304.8, 228.6) },
                    { "ARCH B", Tuple.Create(457.2, 304.8) },
                    { "ARCH C", Tuple.Create(609.6, 457.2) },
                    { "ARCH D", Tuple.Create(914.4, 609.6) },
                    { "ARCH E", Tuple.Create(1219.2, 914.4) },
                    { "ARCH E1", Tuple.Create(1066.8, 762.0) },
                };

                // Try to find the Revit paper size in standard sizes
                Tuple<double, double> targetSize = null;
                foreach (var kvp in standardSizes)
                {
                    if (normalizedRevit.Contains(kvp.Key.ToUpperInvariant()))
                    {
                        targetSize = kvp.Value;
                        break;
                    }
                }

                // Parse dimension-based strings like "42 x 30 mm", "42x30", "841 x 594 mm"
                // Revit exports paper sizes using the title block's sheet size parameter which may be
                // in inches (e.g. "42 x 30 mm" label on an Arch E1 title block = 42"x30" = 1066.8x762mm)
                // or in mm (e.g. "841 x 594 mm" = A1).
                // Heuristic: common Arch inch sizes are small integers (8-48); common ISO mm sizes are
                // large integers (148-1682). Threshold: if both dims <= 60, treat as inches.
                if (targetSize == null)
                {
                    var dimensionMatch = System.Text.RegularExpressions.Regex.Match(
                        normalizedRevit, @"(\d+\.?\d*)\s*[Xx]\s*(\d+\.?\d*)");
                    if (dimensionMatch.Success)
                    {
                        double dim1 = double.Parse(dimensionMatch.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                        double dim2 = double.Parse(dimensionMatch.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);

                        // Explicit unit markers
                        bool explicitInches = normalizedRevit.Contains("INCH") || normalizedRevit.Contains("\"");
                        bool explicitMm = normalizedRevit.Contains(" MM") || normalizedRevit.EndsWith("MM");

                        bool isInches;
                        if (explicitInches)
                            isInches = true;
                        else if (explicitMm && (dim1 > 60.0 || dim2 > 60.0))
                            isInches = false;  // large mm value explicitly labelled mm
                        else if (dim1 <= 60.0 && dim2 <= 60.0)
                            isInches = true;   // small dims with no unit = Arch inches (42x30, 36x24, etc.)
                        else
                            isInches = false;  // large dims = already mm (841x594, etc.)

                        if (isInches)
                        {
                            targetSize = Tuple.Create(dim1 * 25.4, dim2 * 25.4);
                            AcadLogger.LogInfo($"[PaperSize] Parsed inches: {dim1}\"x{dim2}\" = {targetSize.Item1:F1}x{targetSize.Item2:F1}mm");
                        }
                        else
                        {
                            targetSize = Tuple.Create(dim1, dim2);
                            AcadLogger.LogInfo($"[PaperSize] Parsed mm: {dim1}x{dim2}mm");
                        }
                    }
                }

                if (targetSize == null)
                {
                    AcadLogger.LogDebug($"[PaperSize] Revit paper size '{revitPaperSize}' not recognized as standard size");
                    return null;
                }

                double targetW = targetSize.Item1;
                double targetH = targetSize.Item2;

                // Find the best matching media from cache
                string bestMatch = null;
                double bestScore = double.MaxValue;
                const double tolerance = 5.0; // mm

                foreach (var kvp in _mediaSizeCache)
                {
                    double mw = kvp.Value.Item1;
                    double mh = kvp.Value.Item2;

                    // Check both orientations
                    double d1 = Math.Abs(mw - targetW) + Math.Abs(mh - targetH);
                    double d2 = Math.Abs(mw - targetH) + Math.Abs(mh - targetW);
                    double score = Math.Min(d1, d2);

                    if (score < bestScore)
                    {
                        bestScore = score;
                        bestMatch = kvp.Key;
                    }

                    if (score <= tolerance)
                        break;
                }

                if (bestMatch != null && bestScore <= tolerance * 2)
                {
                    AcadLogger.LogInfo($"[PaperSize] Revit '{revitPaperSize}' ({targetW}x{targetH}mm) -> matched media '{bestMatch}' (score={bestScore:F1})");
                    return bestMatch;
                }

                AcadLogger.LogDebug($"[PaperSize] No close media match for Revit '{revitPaperSize}' ({targetW}x{targetH}mm), best='{bestMatch}' score={bestScore:F1}");
                return null;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"[PaperSize] Error matching Revit size '{revitPaperSize}': {ex.Message}");
                return null;
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(originalMedia))
                {
                    try { psv.SetCanonicalMediaName(layout, originalMedia); } catch { }
                }
            }
        }

        private string FindLargerMedia(double requiredWidth, double requiredHeight)
        {
            string bestLarger = null;
            double bestLargerArea = double.MaxValue;

            foreach (var kvp in _mediaSizeCache)
            {
                double mw = kvp.Value.Item1;
                double mh = kvp.Value.Item2;

                bool fitsNormal = mw >= requiredWidth - 2.0 && mh >= requiredHeight - 2.0;
                bool fitsRotated = mw >= requiredHeight - 2.0 && mh >= requiredWidth - 2.0;

                if (fitsNormal || fitsRotated)
                {
                    double area = mw * mh;
                    if (area < bestLargerArea)
                    {
                        bestLargerArea = area;
                        bestLarger = kvp.Key;
                    }
                }
            }

            return bestLarger;
        }

        private void TryApplyCorrectedPaperSize(
            PlotSettingsValidator psv,
            Layout layout,
            double requiredWidth,
            double requiredHeight,
            string layoutName,
            string phase)
        {
            try
            {
                // Check cache first — same required size always yields same best media
                string cacheKey = $"{requiredWidth:F0}x{requiredHeight:F0}";
                if (_correctedPaperCache.TryGetValue(cacheKey, out var cached))
                {
                    psv.SetCanonicalMediaName(layout, cached.Item1);
                    psv.SetPlotRotation(layout, cached.Item2 ? PlotRotation.Degrees090 : PlotRotation.Degrees000);
                    AcadLogger.LogInfo(
                        $"{phase}: Corrected paper size for '{layoutName}' => " +
                        $"media='{cached.Item1}', rotated={cached.Item2} (from cache)");
                    return;
                }

                var mediaNames = psv.GetCanonicalMediaNameList(layout);
                string bestMatch = null;
                double bestScore = double.MaxValue;
                bool bestIsRotated = false;

                foreach (string mediaName in mediaNames)
                {
                    if (string.IsNullOrWhiteSpace(mediaName))
                        continue;

                    // Use cache to avoid repeated SetCanonicalMediaName calls
                    Tuple<double, double> sz;
                    if (!_mediaSizeCache.TryGetValue(mediaName, out sz))
                    {
                        try
                        {
                            psv.SetCanonicalMediaName(layout, mediaName);
                            double mw2 = layout.PlotPaperSize.X;
                            double mh2 = layout.PlotPaperSize.Y;
                            if (mw2 < 1.0 || mh2 < 1.0) continue;
                            sz = Tuple.Create(mw2, mh2);
                            _mediaSizeCache[mediaName] = sz;
                        }
                        catch { continue; }
                    }

                    double mw = sz.Item1;
                    double mh = sz.Item2;
                    double dNormal = Math.Abs(mw - requiredWidth) + Math.Abs(mh - requiredHeight);
                    double dRotated = Math.Abs(mw - requiredHeight) + Math.Abs(mh - requiredWidth);

                    if (dNormal <= dRotated && dNormal < bestScore)
                    {
                        bestScore = dNormal;
                        bestMatch = mediaName;
                        bestIsRotated = false;
                    }
                    else if (dRotated < dNormal && dRotated < bestScore)
                    {
                        bestScore = dRotated;
                        bestMatch = mediaName;
                        bestIsRotated = true;
                    }

                    // Early exit if exact match found
                    if (bestScore <= 2.0)
                        break;
                }

                if (bestMatch != null)
                {
                    _correctedPaperCache[cacheKey] = Tuple.Create(bestMatch, bestIsRotated);
                    psv.SetCanonicalMediaName(layout, bestMatch);
                    psv.SetPlotRotation(layout, bestIsRotated ? PlotRotation.Degrees090 : PlotRotation.Degrees000);
                    AcadLogger.LogInfo(
                        $"{phase}: Corrected paper size for '{layoutName}' => " +
                        $"media='{bestMatch}', paper=({layout.PlotPaperSize.X:F0}x{layout.PlotPaperSize.Y:F0}), " +
                        $"rotated={bestIsRotated}, score={bestScore:F1}");
                }
                else
                {
                    AcadLogger.LogWarning(
                        $"{phase}: No suitable media found for '{layoutName}' " +
                        $"required=({requiredWidth:F0}x{requiredHeight:F0})");
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"{phase}: Paper size correction failed for '{layoutName}': {ex.Message}");
            }
        }

        private Extents3d GetPaperBackgroundExtents(Layout layout, BlockTableRecord paperSpace, Transaction tr, out int extentEntities)
        {
            var entityExtents = GetExtentsExcludingLayer(paperSpace, tr, PaperBackgroundLayerName, out extentEntities);
            bool hasEntityExtents = IsUsableExtents(entityExtents);
            bool hasPlotSize = layout.PlotPaperSize.X > 1.0 && layout.PlotPaperSize.Y > 1.0;

            double minX = hasEntityExtents ? entityExtents.MinPoint.X : 0.0;
            double minY = hasEntityExtents ? entityExtents.MinPoint.Y : 0.0;
            double maxX = hasEntityExtents ? entityExtents.MaxPoint.X : 0.0;
            double maxY = hasEntityExtents ? entityExtents.MaxPoint.Y : 0.0;

            if (hasPlotSize)
            {
                minX = Math.Min(minX, 0.0);
                minY = Math.Min(minY, 0.0);
                maxX = Math.Max(maxX, layout.PlotPaperSize.X);
                maxY = Math.Max(maxY, layout.PlotPaperSize.Y);
            }

            if (!hasEntityExtents && !hasPlotSize)
            {
                minX = 0.0;
                minY = 0.0;
                maxX = PaperBackgroundFallbackWidth;
                maxY = PaperBackgroundFallbackHeight;
            }

            if (maxX - minX < 1.0)
                maxX = minX + PaperBackgroundFallbackWidth;

            if (maxY - minY < 1.0)
                maxY = minY + PaperBackgroundFallbackHeight;

            return new Extents3d(new Point3d(minX, minY, 0.0), new Point3d(maxX, maxY, 0.0));
        }

        private Extents3d GetExtentsExcludingLayer(BlockTableRecord btr, Transaction tr, string excludedLayer, out int extentsEntityCount)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            extentsEntityCount = 0;

            foreach (ObjectId id in btr)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased)
                        continue;

                    if (string.Equals(ent.Layer, excludedLayer, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var ext = ent.GeometricExtents;
                    minX = Math.Min(minX, ext.MinPoint.X);
                    minY = Math.Min(minY, ext.MinPoint.Y);
                    maxX = Math.Max(maxX, ext.MaxPoint.X);
                    maxY = Math.Max(maxY, ext.MaxPoint.Y);
                    extentsEntityCount++;
                }
                catch
                {
                }
            }

            if (minX == double.MaxValue)
                return new Extents3d(Point3d.Origin, Point3d.Origin);

            return new Extents3d(new Point3d(minX, minY, 0.0), new Point3d(maxX, maxY, 0.0));
        }

        private ObjectIdCollection CreateWhiteBackgroundHatch(Database db, Transaction tr, BlockTableRecord paperSpace, Extents3d extents)
        {
            var ids = new ObjectIdCollection();
            var white = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

            var boundary = new Polyline(4);
            boundary.SetDatabaseDefaults(db);
            boundary.Layer = PaperBackgroundLayerName;
            boundary.Color = white;
            boundary.AddVertexAt(0, new Point2d(extents.MinPoint.X, extents.MinPoint.Y), 0.0, 0.0, 0.0);
            boundary.AddVertexAt(1, new Point2d(extents.MaxPoint.X, extents.MinPoint.Y), 0.0, 0.0, 0.0);
            boundary.AddVertexAt(2, new Point2d(extents.MaxPoint.X, extents.MaxPoint.Y), 0.0, 0.0, 0.0);
            boundary.AddVertexAt(3, new Point2d(extents.MinPoint.X, extents.MaxPoint.Y), 0.0, 0.0, 0.0);
            boundary.Closed = true;

            ObjectId boundaryId = paperSpace.AppendEntity(boundary);
            tr.AddNewlyCreatedDBObject(boundary, true);
            ids.Add(boundaryId);

            var hatch = new Hatch();
            hatch.SetDatabaseDefaults(db);
            hatch.Layer = PaperBackgroundLayerName;
            hatch.Color = white;

            ObjectId hatchId = paperSpace.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);

            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            hatch.Associative = true;

            var loopIds = new ObjectIdCollection { boundaryId };
            hatch.AppendLoop(HatchLoopTypes.External, loopIds);
            hatch.EvaluateHatch(true);
            ids.Add(hatchId);

            return ids;
        }

        private bool MoveEntitiesToBottom(BlockTableRecord paperSpace, Transaction tr, ObjectIdCollection ids)
        {
            try
            {
                if (paperSpace.DrawOrderTableId.IsNull || ids == null || ids.Count == 0)
                    return false;

                var drawOrder = (DrawOrderTable)tr.GetObject(paperSpace.DrawOrderTableId, OpenMode.ForWrite);
                drawOrder.MoveToBottom(ids);
                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"PAPER BACKGROUND: failed to move paper background to bottom: {ex.Message}");
                return false;
            }
        }

        private void EnsureLayoutViewportsOnTop(Transaction tr, BlockTableRecord paperSpace, string layoutName)
        {
            try
            {
                if (paperSpace == null || paperSpace.DrawOrderTableId.IsNull)
                    return;

                var viewportIds = new ObjectIdCollection();
                foreach (ObjectId id in paperSpace)
                {
                    try
                    {
                        var vp = tr.GetObject(id, OpenMode.ForRead, false) as Viewport;
                        if (vp == null || vp.IsErased)
                            continue;

                        // Keep both model viewports and the default paper viewport above hatches/masks.
                        viewportIds.Add(id);
                    }
                    catch
                    {
                    }
                }

                if (viewportIds.Count == 0)
                    return;

                var drawOrder = (DrawOrderTable)tr.GetObject(paperSpace.DrawOrderTableId, OpenMode.ForWrite);
                drawOrder.MoveToTop(viewportIds);
                AcadLogger.LogInfo($"RECREATE draw-order: moved {viewportIds.Count} viewport(s) to top for '{layoutName}'");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"RECREATE draw-order failed for '{layoutName}': {ex.Message}");
            }
        }

        private bool IsUsableExtents(Extents3d extents)
        {
            return Math.Abs(extents.MaxPoint.X - extents.MinPoint.X) > 1.0 ||
                   Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y) > 1.0;
        }

        private Extents3d GetModelSpaceSheetBounds(
            Layout layout,
            BlockTableRecord paperSpace,
            Transaction tr,
            IReadOnlyList<ViewportInfo> viewports,
            out int extentEntities)
        {
            string cacheKey = layout != null ? GetExtentsCacheKey(null, layout.LayoutName + "_sheetbounds") : null;
            var bounds = GetPaperEntityExtentsExcludingViewports(paperSpace, tr, out extentEntities, cacheKey);
            bool hasBounds = IsUsableExtents(bounds);

            if (viewports != null)
            {
                foreach (var vp in viewports)
                {
                    var viewportPaperWindow = GetViewportPaperExtents(vp);
                    if (IsUsableExtents(viewportPaperWindow))
                    {
                        bounds = hasBounds ? CombineExtents(bounds, viewportPaperWindow) : viewportPaperWindow;
                        hasBounds = true;
                    }
                }
            }

            if (layout != null && layout.PlotPaperSize.X > 1.0 && layout.PlotPaperSize.Y > 1.0)
            {
                var plotBounds = new Extents3d(
                    new Point3d(0.0, 0.0, 0.0),
                    new Point3d(layout.PlotPaperSize.X, layout.PlotPaperSize.Y, 0.0));
                bounds = hasBounds ? CombineExtents(bounds, plotBounds) : plotBounds;
                hasBounds = true;
            }

            if (!hasBounds)
            {
                bounds = new Extents3d(
                    new Point3d(0.0, 0.0, 0.0),
                    new Point3d(PaperBackgroundFallbackWidth, PaperBackgroundFallbackHeight, 0.0));
            }

            return bounds;
        }

        private Extents3d GetPaperEntityExtentsExcludingViewports(BlockTableRecord paperSpace, Transaction tr, out int extentsEntityCount, string cacheKey = null)
        {
            if (!string.IsNullOrEmpty(cacheKey) && TryGetExtentsCache(cacheKey, out var cached))
            {
                extentsEntityCount = cached.ExtentsEntityCount;
                return cached.Extents;
            }

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            extentsEntityCount = 0;
            var sheetCandidates = new List<Extents3d>();
            var titleBlockBounds = new List<Extents3d>();

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased || ent is Viewport)
                        continue;

                    if (string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (TryGetTitleSheetBounds(ent, out var frameBounds))
                        sheetCandidates.Add(frameBounds);

                    if (TryGetTitleBlockOnlyBounds(ent, tr, out var tbBounds))
                        titleBlockBounds.Add(tbBounds);

                    var ext = ent.GeometricExtents;
                    minX = Math.Min(minX, ext.MinPoint.X);
                    minY = Math.Min(minY, ext.MinPoint.Y);
                    maxX = Math.Max(maxX, ext.MaxPoint.X);
                    maxY = Math.Max(maxY, ext.MaxPoint.Y);
                    extentsEntityCount++;
                }
                catch
                {
                }
            }

            Extents3d result;

            if (titleBlockBounds.Count > 0)
            {
                var bestTitleBlock = SelectBestTitleSheetCandidate(titleBlockBounds);
                AcadLogger.LogInfo(
                    $"PAPER BOUNDS: using title block bounds from {titleBlockBounds.Count} candidates, " +
                    $"result={FormatExtents(bestTitleBlock)}");
                result = bestTitleBlock;
            }
            else if (sheetCandidates.Count > 0)
            {
                var titleBounds = SelectBestTitleSheetCandidate(sheetCandidates);

                double expandLeft  = Math.Max(0, titleBounds.MinPoint.X - minX);
                double expandRight = Math.Max(0, maxX - titleBounds.MaxPoint.X);
                double expandDown  = Math.Max(0, titleBounds.MinPoint.Y - minY);
                double expandUp    = Math.Max(0, maxY - titleBounds.MaxPoint.Y);

                if (expandLeft > 1.0 || expandRight > 1.0 || expandDown > 1.0 || expandUp > 1.0)
                {
                    result = new Extents3d(
                        new Point3d(titleBounds.MinPoint.X - expandLeft,
                                    titleBounds.MinPoint.Y - expandDown, 0.0),
                        new Point3d(titleBounds.MaxPoint.X + expandRight,
                                    titleBounds.MaxPoint.Y + expandUp, 0.0));
                    AcadLogger.LogInfo(
                        $"PAPER BOUNDS: expanded title sheet to include extra content, " +
                        $"title={FormatExtents(titleBounds)}, " +
                        $"expand=(L={expandLeft:F1},R={expandRight:F1},D={expandDown:F1},U={expandUp:F1}), " +
                        $"result={FormatExtents(result)}");
                }
                else
                {
                    result = titleBounds;
                }
            }
            else
            {
                result = minX == double.MaxValue
                    ? new Extents3d(Point3d.Origin, Point3d.Origin)
                    : new Extents3d(new Point3d(minX, minY, 0.0), new Point3d(maxX, maxY, 0.0));
            }

            if (!string.IsNullOrEmpty(cacheKey))
            {
                SetExtentsCache(cacheKey, new ExtentsCacheEntry
                {
                    Extents = result,
                    EntityCount = 0,
                    ExtentsEntityCount = extentsEntityCount,
                    CachedAt = DateTime.Now,
                    IsModelSpace = false
                });
            }

            return result;
        }

        private bool TryGetTitleBlockOnlyBounds(Entity ent, Transaction tr, out Extents3d bounds)
        {
            bounds = new Extents3d(Point3d.Origin, Point3d.Origin);

            try
            {
                if (ent is BlockReference br)
                {
                    string blockName = br.Name ?? "";
                    bool isTitleBlock = TitleBlockKeywords.Any(kw =>
                        blockName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        blockName.IndexOf("TB", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        blockName.IndexOf("BORDER", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        blockName.IndexOf("KHUNG", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isTitleBlock)
                    {
                        bounds = br.GeometricExtents;
                        if (IsReasonableSheetBounds(bounds))
                            return true;
                    }
                }

                if (ent is Polyline lw && lw.Closed && lw.NumberOfVertices >= 4)
                {
                    bounds = lw.GeometricExtents;
                    double w = bounds.MaxPoint.X - bounds.MinPoint.X;
                    double h = bounds.MaxPoint.Y - bounds.MinPoint.Y;

                    bool isBorderLayer = string.Equals(ent.Layer, "DEFPOINTS", StringComparison.OrdinalIgnoreCase) ||
                        ent.Layer.IndexOf("BORDER", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        ent.Layer.IndexOf("TITLE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        ent.Layer.IndexOf("KHUNG", StringComparison.OrdinalIgnoreCase) >= 0;

                    if ((w >= 800.0 && w <= 1300.0 && h >= 500.0 && h <= 950.0) || isBorderLayer)
                    {
                        if (IsReasonableSheetBounds(bounds))
                            return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private bool TryGetTitleSheetBounds(Entity ent, out Extents3d bounds)
        {
            bounds = new Extents3d(Point3d.Origin, Point3d.Origin);

            try
            {
                if (ent is Polyline lw)
                {
                    if (!lw.Closed || lw.NumberOfVertices < 4)
                        return false;

                    bounds = lw.GeometricExtents;
                    if (!IsReasonableSheetBounds(bounds))
                        return false;

                    double w = bounds.MaxPoint.X - bounds.MinPoint.X;
                    double h = bounds.MaxPoint.Y - bounds.MinPoint.Y;
                    if (w < 300.0 || h < 200.0)
                        return false;

                    return true;
                }

                if (ent is Polyline2d p2d)
                {
                    if (!p2d.Closed)
                        return false;

                    bounds = p2d.GeometricExtents;
                    return IsReasonableSheetBounds(bounds);
                }
            }
            catch
            {
            }

            return false;
        }

        private bool IsReasonableSheetBounds(Extents3d bounds)
        {
            double w = bounds.MaxPoint.X - bounds.MinPoint.X;
            double h = bounds.MaxPoint.Y - bounds.MinPoint.Y;
            if (w < 100.0 || h < 100.0)
                return false;

            double area = w * h;
            return area >= 25000.0;
        }

        private static readonly HashSet<string> TitleBlockKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "TITLE", "TITLEBLOCK", "TITLE_BLOCK", "SHEET", "SHEETBLOCK", "SHEET_BLOCK",
            "BORDER", "FRAME", "TITLEBAR", "KHUDEN", "KHUNG"
        };

        private bool IsTitleBlockReference(BlockReference br, Transaction tr)
        {
            try
            {
                string blockName = br.Name ?? "";
                if (TitleBlockKeywords.Any(kw => blockName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                    return true;

                try
                {
                    var ext = br.GeometricExtents;
                    if (IsReasonableSheetBounds(ext))
                        return true;
                }
                catch { }
            }
            catch { }
            return false;
        }

        private Extents3d SelectBestTitleSheetCandidate(List<Extents3d> candidates)
        {
            var best = candidates[0];
            double bestScore = double.MinValue;

            foreach (var c in candidates)
            {
                double w = c.MaxPoint.X - c.MinPoint.X;
                double h = c.MaxPoint.Y - c.MinPoint.Y;
                if (w <= 1.0 || h <= 1.0)
                    continue;

                double area = w * h;
                double nearOriginPenalty = Math.Abs(c.MinPoint.X) + Math.Abs(c.MinPoint.Y);
                double ratio = w > h ? w / h : h / w;
                double ratioPenalty = Math.Abs(ratio - 1.414);

                double score = area - nearOriginPenalty * 2000.0 - ratioPenalty * 100000.0;
                if (score > bestScore)
                {
                    bestScore = score;
                    best = c;
                }
            }

            return best;
        }

        private int CreateModelSpaceSheetBackground(
            Database db,
            Transaction tr,
            BlockTableRecord modelSpace,
            Extents3d sourceBounds,
            List<ObjectId> placedIds)
        {
            try
            {
                var backgroundIds = CreateWhiteBackgroundHatch(db, tr, modelSpace, sourceBounds);
                foreach (ObjectId id in backgroundIds)
                    placedIds.Add(id);

                bool moved = MoveEntitiesToBottom(modelSpace, tr, backgroundIds);
                AcadLogger.LogInfo($"MODELSPACE background: ids={backgroundIds.Count}, movedToBottom={moved}, bounds={FormatExtents(sourceBounds)}");
                return backgroundIds.Count;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"MODELSPACE background failed: {ex.Message}");
                return 0;
            }
        }

        private int ClonePaperEntitiesToModelSpace(
            Database sourceDb,
            Transaction sourceTrans,
            Database outputDb,
            Transaction outputTrans,
            BlockTableRecord sourcePaperSpace,
            BlockTableRecord outputModelSpace,
            List<ObjectId> placedIds,
            string label)
        {
            var sourceIds = new ObjectIdCollection();

            foreach (ObjectId id in sourcePaperSpace)
            {
                try
                {
                    var ent = sourceTrans.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased || ent is Viewport)
                        continue;

                    if (string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    sourceIds.Add(id);
                }
                catch
                {
                }
            }

            if (sourceIds.Count == 0)
            {
                AcadLogger.LogWarning($"MODELSPACE: '{label}' has no paper-space entities to clone");
                return 0;
            }

            var idMap = new IdMapping();
            sourceDb.WblockCloneObjects(sourceIds, outputModelSpace.ObjectId, idMap, DuplicateRecordCloning.Ignore, false);

            int clonedCount = 0;
            foreach (ObjectId sourceId in sourceIds)
            {
                try
                {
                    if (!idMap.Contains(sourceId))
                        continue;

                    var destId = idMap[sourceId].Value;
                    if (destId.IsNull)
                        continue;

                    var ent = outputTrans.GetObject(destId, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased)
                        continue;

                    placedIds.Add(destId);
                    clonedCount++;
                }
                catch
                {
                }
            }

            AcadLogger.LogInfo($"MODELSPACE: '{label}' cloned paper-space entities={clonedCount}/{sourceIds.Count}");
            return clonedCount;
        }

        private int TransformEntities(Transaction tr, IEnumerable<ObjectId> ids, Matrix3d transform)
        {
            int movedCount = 0;

            foreach (ObjectId id in ids)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                    if (ent == null || ent.IsErased)
                        continue;

                    ent.TransformBy(transform);
                    movedCount++;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"MODELSPACE transform failed for id={id}: {ex.Message}");
                }
            }

            return movedCount;
        }

        private int MoveModelSpaceBackgroundsToBottom(BlockTableRecord modelSpace, Transaction tr)
        {
            var backgroundIds = new ObjectIdCollection();

            foreach (ObjectId id in modelSpace)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent != null &&
                        !ent.IsErased &&
                        string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        backgroundIds.Add(id);
                    }
                }
                catch
                {
                }
            }

            if (backgroundIds.Count == 0)
                return 0;

            MoveEntitiesToBottom(modelSpace, tr, backgroundIds);
            return backgroundIds.Count;
        }

        private int RegenerateModelSpace(Database db)
        {
            int entityCount = 0;
            int extentsEntityCount = 0;
            Extents3d extents = new Extents3d(Point3d.Origin, Point3d.Origin);
            var previousWorkingDb = HostApplicationServices.WorkingDatabase;

            try
            {
                AcadLogger.LogSection("Regenerating ModelSpace");
                HostApplicationServices.WorkingDatabase = db;

                try
                {
                    db.TileMode = true;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"MODELSPACE REGEN: failed to switch TileMode: {ex.Message}");
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var modelSpace = (BlockTableRecord)tr.GetObject(
                        SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                    foreach (ObjectId id in modelSpace)
                    {
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                            if (ent != null && !ent.IsErased)
                                entityCount++;
                        }
                        catch
                        {
                        }
                    }

                    extents = GetExtents(modelSpace, tr, out extentsEntityCount);
                    tr.Commit();
                }

                try
                {
                    db.UpdateExt(true);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"MODELSPACE REGEN: UpdateExt failed: {ex.Message}");
                }

                AcadLogger.LogInfo(
                    $"MODELSPACE REGEN complete: entities={entityCount}, extentsEntities={extentsEntityCount}, extents={FormatExtents(extents)}");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"RegenerateModelSpace failed: {ex.Message}");
            }
            finally
            {
                try
                {
                    HostApplicationServices.WorkingDatabase = previousWorkingDb;
                }
                catch
                {
                }
            }

            return entityCount;
        }

        private int RegenerateLayouts(Database db, string mode)
        {
            var layoutInfos = new List<LayoutRegenInfo>();
            int regeneratedCount = 0;
            var previousWorkingDb = HostApplicationServices.WorkingDatabase;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    
                    foreach (DBDictionaryEntry entry in layouts)
                    {
                        if (string.Equals(entry.Key, "Model", StringComparison.OrdinalIgnoreCase))
                            continue;

                        try
                        {
                            var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                            var paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                            int paperEntityCount = 0;
                            int viewportCount = 0;

                            foreach (ObjectId id in paperSpace)
                            {
                                try
                                {
                                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                                    if (ent == null || ent.IsErased) continue;
                                    paperEntityCount++;
                                    if (ent is Viewport) viewportCount++;
                                }
                                catch { }
                            }

                            int extentsCount;
                            var layoutCacheKey = GetExtentsCacheKey(null, entry.Key);
                            var paperExtents = GetExtents(paperSpace, tr, out extentsCount, layoutCacheKey);

                            layoutInfos.Add(new LayoutRegenInfo
                            {
                                Name = entry.Key,
                                TabOrder = layout.TabOrder,
                                PaperEntityCount = paperEntityCount,
                                ViewportCount = viewportCount,
                                ExtentsEntityCount = extentsCount,
                                PaperExtents = paperExtents,
                                RequiresRegen = viewportCount > 0
                            });
                        }
                        catch { }
                    }
                    tr.Commit();
                }

                var layoutsToRegen = layoutInfos.OrderBy(x => x.TabOrder).Where(x => x.RequiresRegen).ToList();

                AcadLogger.LogSection($"Regenerating Layouts ({mode})");
                AcadLogger.LogInfo($"In-db diagnostic pass: {layoutsToRegen.Count} layout(s) have viewports");

                HostApplicationServices.WorkingDatabase = db;

                try
                {
                    db.UpdateExt(true);
                }
                catch (System.Exception updateEx)
                {
                    AcadLogger.LogWarning($"REGENERATING LAYOUT: UpdateExt failed: {updateEx.Message}");
                }

                regeneratedCount = layoutsToRegen.Count;
                AcadLogger.LogInfo($"REGENERATING LAYOUT in-db pass complete: {regeneratedCount} layout(s) refreshed");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"RegenerateLayouts failed for {mode}: {ex.Message}");
            }
            finally
            {
                try
                {
                    HostApplicationServices.WorkingDatabase = previousWorkingDb;
                }
                catch { }
            }

            return regeneratedCount;
        }

private void BindXrefsSafe(Database db)
{
    try
    {
        SetXrefPathsToRelative(db);

        try
        {
            db.ResolveXrefs(false, false);
            AcadLogger.LogInfo("Resolved XREF paths before bind");
        }
        catch (System.Exception resolveEx)
        {
            AcadLogger.LogWarning($"ResolveXrefs: {resolveEx.Message}");
        }

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var xrefIds = new ObjectIdCollection();

            foreach (ObjectId btrId in bt)
            {
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                if (btr.IsFromExternalReference)
                    xrefIds.Add(btrId);
            }

            if (xrefIds.Count > 0)
            {
                db.BindXrefs(xrefIds, true);
                AcadLogger.LogInfo($"Bound {xrefIds.Count} XREF(s)");
            }
            else
            {
                AcadLogger.LogInfo("No XREFs to bind");
            }

            tr.Commit();
        }
    }
    catch (System.Exception ex)
    {
        AcadLogger.LogWarning($"BindXrefsSafe: {ex.Message}");
    }
}

private void SetXrefPathsToRelative(Database db)
{
    try
    {
        var hostDwg = db?.Filename;
        var hostDir = string.IsNullOrWhiteSpace(hostDwg) ? null : Path.GetDirectoryName(hostDwg);
        if (string.IsNullOrWhiteSpace(hostDir) || !Directory.Exists(hostDir))
            return;

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            int updated = 0;

            foreach (ObjectId btrId in bt)
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null || !btr.IsFromExternalReference)
                    continue;

                var currentPath = btr.PathName;
                if (string.IsNullOrWhiteSpace(currentPath) || !Path.IsPathRooted(currentPath))
                    continue;

                string relativePath = TryGetRelativePath(hostDir, currentPath);

                if (string.IsNullOrWhiteSpace(relativePath) || relativePath == currentPath)
                    continue;

                try
                {
                    var btrWrite = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                    btrWrite.PathName = relativePath;
                    updated++;
                }
                catch (System.Exception setEx)
                {
                    AcadLogger.LogWarning($"SetXrefPathsToRelative: cannot update '{btr.Name}': {setEx.Message}");
                }
            }

            tr.Commit();

            if (updated > 0)
                AcadLogger.LogInfo($"XREF change path type: Make Relative applied to {updated} reference(s)");
        }
    }
    catch (System.Exception ex)
    {
        AcadLogger.LogWarning($"SetXrefPathsToRelative: {ex.Message}");
    }
}

private string TryGetRelativePath(string baseDir, string targetPath)
{
    try
    {
        if (string.IsNullOrWhiteSpace(baseDir) || string.IsNullOrWhiteSpace(targetPath))
            return targetPath;

        var baseUri = new Uri(AppendDirectorySeparator(baseDir));
        var targetUri = new Uri(targetPath);

        if (!string.Equals(baseUri.Scheme, targetUri.Scheme, StringComparison.OrdinalIgnoreCase))
            return targetPath;

        var relativeUri = baseUri.MakeRelativeUri(targetUri);
        var relativePath = Uri.UnescapeDataString(relativeUri.ToString()).Replace('/', Path.DirectorySeparatorChar);
        return string.IsNullOrWhiteSpace(relativePath) ? targetPath : relativePath;
    }
    catch
    {
        return targetPath;
    }
}

private static string AppendDirectorySeparator(string path)
{
    if (string.IsNullOrEmpty(path))
        return path;

    char lastChar = path[path.Length - 1];
    if (lastChar != Path.DirectorySeparatorChar && lastChar != Path.AltDirectorySeparatorChar)
        return path + Path.DirectorySeparatorChar;

    return path;
}

private void RenameBlocksInDb(Database db, Transaction trans, string prefix)
{
    var bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
    int renamedCount = 0;
    int skippedCount = 0;

    foreach (ObjectId btrId in bt)
    {
        var btr = (BlockTableRecord)trans.GetObject(btrId, OpenMode.ForRead);
        if (btr.IsLayout || btr.IsAnonymous || btr.IsFromExternalReference)
        {
            skippedCount++;
            continue;
        }

        string oldName = btr.Name;
        if (oldName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            skippedCount++;
            continue;
        }

        string newName = prefix + oldName;

        var btrWrite = (BlockTableRecord)trans.GetObject(btrId, OpenMode.ForWrite);
        btrWrite.Name = newName;
        renamedCount++;
    }

    AcadLogger.LogInfo($"Renamed {renamedCount} block definitions with prefix '{prefix}' (skipped {skippedCount})");
}

private Layout GetSourceLayout(Database db, Transaction trans, string desiredLayoutName)
{
    try
    {
        var layouts = (DBDictionary)trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
        if (!string.IsNullOrEmpty(desiredLayoutName))
        {
            if (layouts.Contains(desiredLayoutName))
            {
                var layoutId = layouts.GetAt(desiredLayoutName);
                return (Layout)trans.GetObject(layoutId, OpenMode.ForRead);
            }
        }
        foreach (var entry in layouts)
        {
            if (entry.Key == "Model") continue;
            return (Layout)trans.GetObject(entry.Value, OpenMode.ForRead);
        }
    }
    catch (System.Exception ex)
    {
        AcadLogger.LogWarning($"GetSourceLayout: {ex.Message}");
    }
    return null;
}


        private int BakeModelViewsToPaperSpace(
            Database sourceDb,
            Transaction sourceTrans,
            Database outputDb,
            Transaction outputTrans,
            BlockTableRecord destPaperSpace,
            IReadOnlyList<ViewportInfo> sourceViewports,
            string layoutName,
            List<ObjectId> transformedDestIds = null)
        {
            if (sourceViewports == null || sourceViewports.Count == 0)
            {
                AcadLogger.LogWarning($"BAKE: '{layoutName}' has no usable source viewport(s)");
                return 0;
            }

            int totalTransformed = 0;

            try
            {
                var sourceModelSpace = (BlockTableRecord)sourceTrans.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(sourceDb), OpenMode.ForRead);

                var modelEntityExtents = BuildModelEntityExtentsCache(sourceTrans, sourceModelSpace, layoutName, out int modelEntityCount, out int modelNoExtentsCount);

                AcadLogger.LogInfo(
                    $"BAKE: Start '{layoutName}' sourceVpCount={sourceViewports.Count}, " +
                    $"sameDb={object.ReferenceEquals(sourceDb, outputDb)}, destPaperSpace={destPaperSpace.ObjectId}, " +
                    $"cachedEntities={modelEntityCount}, noExtents={modelNoExtentsCount}");

                for (int i = 0; i < sourceViewports.Count; i++)
                {
                    var vp = sourceViewports[i];
                    if (vp.Width <= 0.0 || vp.Height <= 0.0 || vp.ViewHeight <= 0.0 || vp.CustomScale <= 0.0)
                    {
                        AcadLogger.LogWarning(
                            $"BAKE: Skip invalid viewport '{layoutName}' index={i + 1} " +
                            $"paperSize=({vp.Width:F2},{vp.Height:F2}) viewHeight={vp.ViewHeight:F2} scale={vp.CustomScale:F8}");
                        continue;
                    }

                    var visibleExtents = GetViewportViewExtents(vp);
                    double viewWidth = visibleExtents.MaxPoint.X - visibleExtents.MinPoint.X;
                    double viewHeight = visibleExtents.MaxPoint.Y - visibleExtents.MinPoint.Y;
                    double selectionMargin = Math.Max(1.0, Math.Min(Math.Abs(viewWidth), Math.Abs(viewHeight)) * 0.01);
                    var searchExtents = ExpandExtents(visibleExtents, selectionMargin);

                    int scannedCount;
                    int noExtentsCount;
                    var idsToBake = CollectModelEntityIdsForViewport(
                        modelEntityExtents, searchExtents, out scannedCount, out noExtentsCount);

                    AcadLogger.LogInfo(
                        $"BAKE: '{layoutName}' viewport {i + 1}/{sourceViewports.Count} " +
                        $"selected={idsToBake.Count}, scanned={scannedCount}, noExtents={noExtentsCount}, " +
                        $"paperCenter={FormatPoint(vp.CenterPoint)} paperSize=({vp.Width:F2},{vp.Height:F2}) " +
                        $"viewCenter={FormatPoint(vp.ViewCenter)} visible={FormatExtents(visibleExtents)} " +
                        $"search={FormatExtents(searchExtents)} scale={vp.CustomScale:F8} twist={vp.TwistAngle:F8}");

                    if (idsToBake.Count == 0)
                        continue;

                    var idMap = new IdMapping();
                    if (object.ReferenceEquals(sourceDb, outputDb))
                    {
                        sourceDb.DeepCloneObjects(idsToBake, destPaperSpace.ObjectId, idMap, false);
                    }
                    else
                    {
                        sourceDb.WblockCloneObjects(
                            idsToBake, destPaperSpace.ObjectId, idMap, DuplicateRecordCloning.Ignore, false);
                    }

                    var transform = GetModelToPaperTransform(vp);
                    var transformedIds = new List<ObjectId>();
                    int clonedCount = 0;
                    int transformedCount = 0;

                    foreach (ObjectId sourceId in idsToBake)
                    {
                        try
                        {
                            if (!idMap.Contains(sourceId))
                                continue;

                            var destId = idMap[sourceId].Value;
                            if (destId.IsNull)
                                continue;

                            clonedCount++;
                            var ent = outputTrans.GetObject(destId, OpenMode.ForWrite, false) as Entity;
                            if (ent == null)
                                continue;

                            ent.TransformBy(transform);
                            transformedIds.Add(destId);
                            transformedDestIds?.Add(destId);
                            transformedCount++;
                        }
                        catch (System.Exception ex)
                        {
                            AcadLogger.LogWarning($"BAKE: Transform failed for '{layoutName}' sourceId={sourceId}: {ex.Message}");
                        }
                    }

                    totalTransformed += transformedCount;
                    int paperExtentsCount;
                    var paperExtents = GetObjectExtents(outputTrans, transformedIds, out paperExtentsCount);
                    var paperWindow = GetViewportPaperExtents(vp);
                    AcadLogger.LogInfo(
                        $"BAKE: '{layoutName}' viewport {i + 1} cloned={clonedCount}, transformed={transformedCount}, " +
                        $"paperExtentsEntities={paperExtentsCount}, paperExtents={FormatExtents(paperExtents)}, " +
                        $"paperWindow={FormatExtents(paperWindow)}");
                }

                AcadLogger.LogInfo($"BAKE: Summary '{layoutName}' transformed={totalTransformed}");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"BAKE failed for '{layoutName}': {ex.Message}");
            }

            return totalTransformed;
        }

        private Extents3d GetObjectExtents(Transaction trans, IEnumerable<ObjectId> ids, out int extentsCount)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            extentsCount = 0;

            foreach (ObjectId id in ids)
            {
                try
                {
                    var ent = trans.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased)
                        continue;

                    var ext = ent.GeometricExtents;
                    minX = Math.Min(minX, ext.MinPoint.X);
                    minY = Math.Min(minY, ext.MinPoint.Y);
                    maxX = Math.Max(maxX, ext.MaxPoint.X);
                    maxY = Math.Max(maxY, ext.MaxPoint.Y);
                    extentsCount++;
                }
                catch
                {
                }
            }

            if (minX == double.MaxValue)
                return new Extents3d(Point3d.Origin, Point3d.Origin);

            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        }

        private ObjectIdCollection CollectModelEntityIdsForViewport(
            IReadOnlyList<ModelEntityExtents> modelEntityExtents,
            Extents3d searchExtents,
            out int scannedCount,
            out int noExtentsCount)
        {
            var ids = new ObjectIdCollection();
            scannedCount = 0;
            noExtentsCount = 0;

            if (modelEntityExtents == null || modelEntityExtents.Count == 0)
                return ids;

            foreach (var item in modelEntityExtents)
            {
                scannedCount++;

                if (ExtentsIntersect(item.Extents, searchExtents))
                    ids.Add(item.Id);
            }

            return ids;
        }

        private List<ModelEntityExtents> BuildModelEntityExtentsCache(
            Transaction trans,
            BlockTableRecord modelSpace,
            string layoutName,
            out int entityCount,
            out int noExtentsCount)
        {
            var result = new List<ModelEntityExtents>();
            entityCount = 0;
            noExtentsCount = 0;

            foreach (ObjectId id in modelSpace)
            {
                entityCount++;
                try
                {
                    var ent = trans.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased)
                        continue;

                    result.Add(new ModelEntityExtents
                    {
                        Id = id,
                        Extents = ent.GeometricExtents
                    });
                }
                catch
                {
                    noExtentsCount++;
                }
            }

            AcadLogger.LogInfo(
                $"BAKE: '{layoutName}' built extents cache entities={entityCount}, cached={result.Count}, noExtents={noExtentsCount}");

            return result;
        }

        private sealed class ModelEntityExtents
        {
            public ObjectId Id { get; set; }
            public Extents3d Extents { get; set; }
        }

        private int EraseAllLayoutViewports(Transaction trans, BlockTableRecord paperSpace, string layoutName)
        {
            var ids = new List<ObjectId>();
            foreach (ObjectId id in paperSpace)
                ids.Add(id);

            int erasedCount = 0;
            int skippedCount = 0;

            foreach (ObjectId id in ids)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForWrite, false) as Viewport;
                    if (vp == null)
                        continue;

                    if (vp.Number == 1)
                    {
                        skippedCount++;
                        AcadLogger.LogInfo($"BAKE: Keep paper viewport #1 for '{layoutName}' handle={vp.Handle}");
                        continue;
                    }

                    AcadLogger.LogInfo(
                        $"BAKE: Erase viewport after paper bake '{layoutName}' VP#{vp.Number} handle={vp.Handle} " +
                        $"paperCenter={FormatPoint(vp.CenterPoint)} paperSize=({vp.Width:F2},{vp.Height:F2}) " +
                        $"viewCenter={FormatPoint(vp.ViewCenter)} scale={vp.CustomScale:F8}");
                    vp.Erase();
                    erasedCount++;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"BAKE: Failed to erase viewport for '{layoutName}' id={id}: {ex.Message}");
                }
            }

            AcadLogger.LogInfo($"BAKE: Erased viewports for '{layoutName}' erased={erasedCount}, skipped={skippedCount}");
            return erasedCount;
        }

        private int EraseAllLayoutViewportsForRecreate(
            Transaction trans,
            BlockTableRecord paperSpace,
            string layoutName)
        {
            var ids = new List<ObjectId>();
            foreach (ObjectId id in paperSpace)
                ids.Add(id);

            int erasedCount = 0;
            int skippedCount = 0;

            foreach (ObjectId id in ids)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForWrite, false) as Viewport;
                    if (vp == null || vp.IsErased)
                        continue;

                    if (!IsRawModelViewport(vp))
                    {
                        skippedCount++;
                        continue;
                    }

                    AcadLogger.LogInfo(
                        $"RECREATE erase viewport '{layoutName}' handle={vp.Handle} " +
                        $"paperCenter={FormatPoint(vp.CenterPoint)} " +
                        $"paperSize=({vp.Width:F2},{vp.Height:F2}) " +
                        $"viewCenter={FormatPoint(vp.ViewCenter)} " +
                        $"scale={vp.CustomScale:F8}");

                    vp.Erase();
                    erasedCount++;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning(
                        $"RECREATE failed to erase viewport for '{layoutName}' id={id}: {ex.Message}");
                }
            }

            AcadLogger.LogInfo(
                $"RECREATE erased viewports for '{layoutName}' erased={erasedCount}, skipped={skippedCount}");

            return erasedCount;
        }

        private Matrix3d GetModelToPaperTransform(ViewportInfo vp)
        {
            double scale = vp.CustomScale > 0.0 ? vp.CustomScale : 1.0;

            var moveModelToOrigin = Matrix3d.Displacement(
                new Vector3d(-vp.ViewCenter.X, -vp.ViewCenter.Y, 0.0));
            var scaleToPaper = Matrix3d.Scaling(scale, Point3d.Origin);
            var rotateToPaper = Math.Abs(vp.TwistAngle) > 1e-10
                ? Matrix3d.Rotation(-vp.TwistAngle, Vector3d.ZAxis, Point3d.Origin)
                : Matrix3d.Identity;
            var moveToViewportCenter = Matrix3d.Displacement(
                new Vector3d(vp.CenterPoint.X, vp.CenterPoint.Y, 0.0));

            return moveToViewportCenter * rotateToPaper * scaleToPaper * moveModelToOrigin;
        }

        private Extents3d GetViewportViewExtents(ViewportInfo vp)
        {
            if (vp == null || vp.Width <= 0.0 || vp.Height <= 0.0 || vp.ViewHeight <= 0.0)
                return new Extents3d(
                    new Point3d(vp?.ViewCenter.X ?? 0.0, vp?.ViewCenter.Y ?? 0.0, 0.0),
                    new Point3d(vp?.ViewCenter.X ?? 0.0, vp?.ViewCenter.Y ?? 0.0, 0.0));

            double viewWidth = vp.ViewHeight * (vp.Width / vp.Height);
            double halfWidth = viewWidth / 2.0;
            double halfHeight = vp.ViewHeight / 2.0;

            return new Extents3d(
                new Point3d(vp.ViewCenter.X - halfWidth, vp.ViewCenter.Y - halfHeight, 0),
                new Point3d(vp.ViewCenter.X + halfWidth, vp.ViewCenter.Y + halfHeight, 0));
        }

        private Extents3d GetViewportPaperExtents(ViewportInfo vp)
        {
            if (vp == null)
                return new Extents3d(Point3d.Origin, Point3d.Origin);

            double halfWidth = vp.Width / 2.0;
            double halfHeight = vp.Height / 2.0;

            return new Extents3d(
                new Point3d(vp.CenterPoint.X - halfWidth, vp.CenterPoint.Y - halfHeight, 0),
                new Point3d(vp.CenterPoint.X + halfWidth, vp.CenterPoint.Y + halfHeight, 0));
        }

        private Extents3d ExpandExtents(Extents3d extents, double margin)
        {
            return new Extents3d(
                new Point3d(extents.MinPoint.X - margin, extents.MinPoint.Y - margin, extents.MinPoint.Z - margin),
                new Point3d(extents.MaxPoint.X + margin, extents.MaxPoint.Y + margin, extents.MaxPoint.Z + margin));
        }

        private bool ExtentsIntersect(Extents3d a, Extents3d b)
        {
            return a.MinPoint.X <= b.MaxPoint.X &&
                a.MaxPoint.X >= b.MinPoint.X &&
                a.MinPoint.Y <= b.MaxPoint.Y &&
                a.MaxPoint.Y >= b.MinPoint.Y;
        }

        private string GetSafeAutoCadLayoutName(string requestedName, HashSet<string> usedLayoutNames)
        {
            var safeName = SanitizeAutoCadLayoutName(requestedName);
            var uniqueName = MakeUniqueLayoutName(safeName, usedLayoutNames);

            if (!string.Equals(requestedName, uniqueName, StringComparison.Ordinal))
            {
                AcadLogger.LogWarning(
                    $"Layout name adjusted for AutoCAD: '{requestedName}' -> '{uniqueName}'");
            }

            return uniqueName;
        }

        private string SanitizeAutoCadLayoutName(string requestedName)
        {
            if (string.IsNullOrWhiteSpace(requestedName))
                return "Layout";

            var invalidChars = new HashSet<char>("<>/\\\":;?*|=,&".ToCharArray());
            var chars = requestedName
                .Trim()
                .Select(c => invalidChars.Contains(c) || char.IsControl(c) ? ' ' : c)
                .ToArray();

            var safeName = new string(chars).Trim();
            while (safeName.Contains("  "))
                safeName = safeName.Replace("  ", " ");

            if (string.IsNullOrWhiteSpace(safeName))
                safeName = "Layout";

            if (safeName.Length > AutoCadLayoutNameMaxLength)
                safeName = safeName.Substring(0, AutoCadLayoutNameMaxLength).TrimEnd();

            return safeName;
        }

        private string MakeUniqueLayoutName(string baseName, HashSet<string> usedLayoutNames)
        {
            if (usedLayoutNames == null)
                return baseName;

            var uniqueName = baseName;
            int suffix = 2;

            while (usedLayoutNames.Contains(uniqueName))
            {
                var suffixText = $" ({suffix})";
                var prefixLength = Math.Max(1, AutoCadLayoutNameMaxLength - suffixText.Length);
                var prefix = baseName.Length > prefixLength
                    ? baseName.Substring(0, prefixLength).TrimEnd()
                    : baseName;
                uniqueName = prefix + suffixText;
                suffix++;
            }

            usedLayoutNames.Add(uniqueName);
            return uniqueName;
        }

        private ObjectId CreateNewLayoutInDb(Database outputDb, Transaction outputTrans, string layoutName)
        {
            try
            {
                // L??y LayoutDictionary vA? BlockTable
                var layoutDict = (DBDictionary)outputTrans.GetObject(outputDb.LayoutDictionaryId, OpenMode.ForWrite);
                
                // Ki?_?m tra layout ?`A? t?_"n t?-i ch?oa
                if (layoutDict.Contains(layoutName))
                {
                    AcadLogger.LogWarning($"Layout '{layoutName}' already exists");
                    return ObjectId.Null;
                }
                
                // T?-o BTR m?_>i v?_>i tAYn unique (khA'ng copy t?_r template)
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var candidate = outputTrans.GetObject(entry.Value, OpenMode.ForWrite, false) as Layout;
                    if (candidate == null || candidate.ModelType || !IsDefaultLayoutName(candidate.LayoutName))
                        continue;

                    var candidateBtr = outputTrans.GetObject(candidate.BlockTableRecordId, OpenMode.ForWrite, false) as BlockTableRecord;
                    if (candidateBtr == null || LayoutHasContent(candidateBtr, outputTrans))
                        continue;

                    var existingIds = new List<ObjectId>();
                    foreach (ObjectId id in candidateBtr)
                        existingIds.Add(id);

                    foreach (var id in existingIds)
                    {
                        try
                        {
                            var entity = outputTrans.GetObject(id, OpenMode.ForWrite, false) as Entity;
                            if (entity != null && !entity.IsErased)
                                entity.Erase();
                        }
                        catch
                        {
                        }
                    }

                    var oldName = candidate.LayoutName;
                    candidate.LayoutName = layoutName;
                    candidate.TabOrder = layoutDict.Count - 1;
                    AcadLogger.LogInfo(
                        $"Reused empty default layout '{oldName}' as '{layoutName}' (BTR={candidateBtr.ObjectId}, Layout={candidate.ObjectId})");
                    return candidateBtr.ObjectId;
                }

                var previousWorkingDb = HostApplicationServices.WorkingDatabase;
                try
                {
                    HostApplicationServices.WorkingDatabase = outputDb;
                    var newLayoutId = LayoutManager.Current.CreateLayout(layoutName);
                    var newLayout = (Layout)outputTrans.GetObject(newLayoutId, OpenMode.ForWrite);
                    newLayout.TabOrder = layoutDict.Count - 1;

                    var newBtrId = newLayout.BlockTableRecordId;
                    AcadLogger.LogInfo($"Created new layout '{layoutName}' via LayoutManager (BTR={newBtrId}, Layout={newLayoutId})");
                    return newBtrId;
                }
                finally
                {
                    HostApplicationServices.WorkingDatabase = previousWorkingDb;
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"CreateNewLayoutInDb failed: {ex.Message}");
                return ObjectId.Null;
            }
        }

        private int EnsurePendingScheduleOnlyLayouts(
            Database outputDb,
            List<SourceFileInfo> pendingLayouts,
            List<SourceFileInfo> sourceInfos)
        {
            if (pendingLayouts == null || pendingLayouts.Count == 0)
                return 0;

            int recoveredCount = 0;
            AcadLogger.LogSection("Schedule-only Layout Recovery");
            AcadLogger.LogInfo($"SCHEDULE RECOVERY: pending={pendingLayouts.Count}");

            foreach (var info in pendingLayouts)
            {
                if (info == null || string.IsNullOrWhiteSpace(info.LayoutName))
                    continue;

                if (sourceInfos.Any(x => string.Equals(x.LayoutName, info.LayoutName, StringComparison.OrdinalIgnoreCase)))
                {
                    AcadLogger.LogInfo($"SCHEDULE RECOVERY: '{info.LayoutName}' already exists in source info list");
                    continue;
                }

                if (!EnsurePendingScheduleOnlyLayout(outputDb, info))
                    continue;

                sourceInfos.Add(CloneSourceInfo(info));
                recoveredCount++;
            }

            AcadLogger.LogInfo($"SCHEDULE RECOVERY complete: recovered={recoveredCount}/{pendingLayouts.Count}");
            return recoveredCount;
        }

        private bool EnsurePendingScheduleOnlyLayout(Database outputDb, SourceFileInfo info)
        {
            try
            {
                using (var tr = outputDb.TransactionManager.StartTransaction())
                {
                    var btrId = GetLayoutBlockTableRecordId(outputDb, tr, info.LayoutName);
                    if (btrId.IsNull)
                        btrId = ReuseEmptyDefaultLayout(outputDb, tr, info.LayoutName);

                    if (!btrId.IsNull)
                    {
                        PrepareRecoveredScheduleLayout(outputDb, tr, btrId, info);
                        tr.Commit();
                        AcadLogger.LogInfo($"SCHEDULE RECOVERY: preserved '{info.LayoutName}' via existing/default layout");
                        return true;
                    }

                    tr.Commit();
                }

                if (!CreateLayoutOutsideTransaction(outputDb, info.LayoutName))
                    return false;

                using (var tr = outputDb.TransactionManager.StartTransaction())
                {
                    var btrId = GetLayoutBlockTableRecordId(outputDb, tr, info.LayoutName);
                    if (btrId.IsNull)
                    {
                        AcadLogger.LogError($"SCHEDULE RECOVERY: layout '{info.LayoutName}' was created but cannot be found");
                        tr.Commit();
                        return false;
                    }

                    PrepareRecoveredScheduleLayout(outputDb, tr, btrId, info);
                    tr.Commit();
                }

                AcadLogger.LogInfo($"SCHEDULE RECOVERY: preserved '{info.LayoutName}' via new layout");
                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"SCHEDULE RECOVERY failed for '{info.LayoutName}': {ex.Message}");
                return false;
            }
        }

        private void PrepareRecoveredScheduleLayout(Database outputDb, Transaction tr, ObjectId btrId, SourceFileInfo info)
        {
            var paperSpace = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
            var layout = (Layout)tr.GetObject(paperSpace.LayoutId, OpenMode.ForWrite);

            if (info.PlotSettings != null)
            {
                var savedBtrId = layout.BlockTableRecordId;
                var savedTabOrder = layout.TabOrder;
                ApplyLayoutPlotSettingsSafely(layout, info.PlotSettings, info.LayoutName, savedBtrId, savedTabOrder, "Pending");

                if (layout.PlotPaperSize.X <= 1e-6 || layout.PlotPaperSize.Y <= 1e-6)
                {
                    AcadLogger.LogWarning(
                        $"Pending: PlotPaperSize invalid after apply for '{info.LayoutName}', attempting geometry fallback");

                    try
                    {
                        EnsurePlotSettingsRefreshed(layout);
                    }
                    catch (System.Exception psvEx)
                    {
                        AcadLogger.LogWarning(
                            $"Pending: PlotSettingsValidator refresh failed for '{info.LayoutName}': {psvEx.Message}");
                    }

                    try
                    {
                        layout.CopyFrom(info.PlotSettings);
                    }
                    catch (System.Exception copyEx)
                    {
                        AcadLogger.LogWarning(
                            $"Pending: PlotSettings fallback CopyFrom failed for '{info.LayoutName}': {copyEx.Message}");
                    }
                }
            }

            EnsureLayoutPaperContextFromGeometry(outputDb, tr, layout, paperSpace, info.LayoutName, "Pending");

            if (!LayoutHasContent(paperSpace, tr))
                AddSchedulePlaceholderContent(paperSpace, tr, info.LayoutName, layout);

            AcadLogger.LogInfo(
                $"SCHEDULE RECOVERY: layout='{info.LayoutName}', btr={btrId}, " +
                $"paperSize=({layout.PlotPaperSize.X:F2},{layout.PlotPaperSize.Y:F2})");
        }

        private ObjectId GetLayoutBlockTableRecordId(Database db, Transaction tr, string layoutName)
        {
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            if (!layoutDict.Contains(layoutName))
                return ObjectId.Null;

            var layout = tr.GetObject(layoutDict.GetAt(layoutName), OpenMode.ForRead, false) as Layout;
            if (layout == null || layout.BlockTableRecordId.IsNull)
                return ObjectId.Null;

            return layout.BlockTableRecordId;
        }

        private bool CreateLayoutOutsideTransaction(Database outputDb, string layoutName)
        {
            var previousWorkingDb = HostApplicationServices.WorkingDatabase;
            try
            {
                HostApplicationServices.WorkingDatabase = outputDb;
                var layoutId = LayoutManager.Current.CreateLayout(layoutName);
                AcadLogger.LogInfo($"SCHEDULE RECOVERY: Created layout '{layoutName}' via LayoutManager outside transaction (Layout={layoutId})");
                return !layoutId.IsNull;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"SCHEDULE RECOVERY: CreateLayout failed for '{layoutName}': {ex.Message}");
                return false;
            }
            finally
            {
                try
                {
                    HostApplicationServices.WorkingDatabase = previousWorkingDb;
                }
                catch
                {
                }
            }
        }

        private SourceFileInfo CloneSourceInfo(SourceFileInfo source)
        {
            return new SourceFileInfo
            {
                FilePath = source.FilePath,
                LayoutName = source.LayoutName,
                MsOffset = source.MsOffset,
                MsExtents = source.MsExtents,
                ModelType = source.ModelType,
                PlotSettings = source.PlotSettings
            };
        }

        private ObjectId ReuseEmptyDefaultLayout(Database outputDb, Transaction outputTrans, string layoutName)
        {
            try
            {
                var layoutDict = (DBDictionary)outputTrans.GetObject(outputDb.LayoutDictionaryId, OpenMode.ForWrite);
                if (layoutDict.Contains(layoutName))
                    return ObjectId.Null;

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var candidate = outputTrans.GetObject(entry.Value, OpenMode.ForWrite, false) as Layout;
                    if (candidate == null || candidate.ModelType || !IsDefaultLayoutName(candidate.LayoutName))
                        continue;

                    var candidateBtr = outputTrans.GetObject(candidate.BlockTableRecordId, OpenMode.ForWrite, false) as BlockTableRecord;
                    if (candidateBtr == null || LayoutHasContent(candidateBtr, outputTrans))
                        continue;

                    var existingIds = new List<ObjectId>();
                    foreach (ObjectId id in candidateBtr)
                        existingIds.Add(id);

                    foreach (var id in existingIds)
                    {
                        try
                        {
                            var entity = outputTrans.GetObject(id, OpenMode.ForWrite, false) as Entity;
                            if (entity != null && !entity.IsErased)
                                entity.Erase();
                        }
                        catch
                        {
                        }
                    }

                    var oldName = candidate.LayoutName;
                    candidate.LayoutName = layoutName;
                    candidate.TabOrder = layoutDict.Count - 1;
                    AcadLogger.LogInfo(
                        $"Reused empty default layout '{oldName}' as '{layoutName}' (BTR={candidateBtr.ObjectId}, Layout={candidate.ObjectId})");
                    return candidateBtr.ObjectId;
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"ReuseEmptyDefaultLayout failed for '{layoutName}': {ex.Message}");
            }

            return ObjectId.Null;
        }

        private void AddSchedulePlaceholderContent(BlockTableRecord paperSpace, Transaction tr, string layoutName, Layout layout)
        {
            try
            {
                var paperWidth = layout != null && layout.PlotPaperSize.X > 1.0 ? layout.PlotPaperSize.X : PaperBackgroundFallbackWidth;
                var paperHeight = layout != null && layout.PlotPaperSize.Y > 1.0 ? layout.PlotPaperSize.Y : PaperBackgroundFallbackHeight;
                var height = Math.Max(2.5, Math.Min(paperWidth, paperHeight) * 0.012);

                var text = new DBText
                {
                    TextString = layoutName,
                    Height = height,
                    Position = new Point3d(Math.Max(10.0, paperWidth * 0.025), Math.Max(10.0, paperHeight * 0.025), 0.0),
                    Layer = "0"
                };

                paperSpace.AppendEntity(text);
                tr.AddNewlyCreatedDBObject(text, true);
                AcadLogger.LogWarning(
                    $"Schedule-only sheet '{layoutName}' had no mergeable DWG geometry; added PaperSpace marker so the layout is preserved.");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"AddSchedulePlaceholderContent failed for '{layoutName}': {ex.Message}");
            }
        }

        private void RenameLayoutInDb(Database db, string oldName, string newName)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite);
                    if (layouts.Contains(oldName))
                    {
                        var layoutId = layouts.GetAt(oldName);
                        var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForWrite);
                        layout.LayoutName = newName;
                        AcadLogger.Log($"[LayoutMerger] Renamed layout '{oldName}' to '{newName}'");
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] RenameLayoutInDb error: {ex.Message}");
            }
        }

        private void CleanupDefaultLayouts(Database db)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite);
                    var layoutsToDelete = new List<string>();

                    foreach (var entry in layouts)
                    {
                        if (entry.Key == "Model")
                            continue;

                        // Ch?_% xA3a layout cA3 tAYn m??c ?`?_<nh (Layout1, Layout2...) vA? BTR r?_-ng
                        // KhA'ng xA3a layout do user ?`??t tAYn
                        bool isDefaultName = entry.Key.StartsWith("Layout", StringComparison.OrdinalIgnoreCase) 
                            && int.TryParse(entry.Key.Substring(6), out _);
                        
                        if (!isDefaultName)
                            continue;

                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        var btrId = layout.BlockTableRecordId;

                        if (btrId == ObjectId.Null)
                        {
                            layoutsToDelete.Add(entry.Key);
                            continue;
                        }

                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.GetEnumerator().MoveNext())
                            layoutsToDelete.Add(entry.Key);
                    }

                    foreach (var name in layoutsToDelete)
                    {
                        layouts.Remove(name);
                        AcadLogger.Log($"[LayoutMerger] Deleted empty default layout '{name}'");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] CleanupDefaultLayouts error: {ex.Message}");
            }
        }

        private BlockTableRecord GetSourcePaperSpace(Database sourceDb, Transaction sourceTrans)
        {
            try
            {
                var layouts = (DBDictionary)sourceTrans.GetObject(sourceDb.LayoutDictionaryId, OpenMode.ForRead);

                foreach (var entry in layouts)
                {
                    if (entry.Key == "Model")
                        continue;

                    var layout = (Layout)sourceTrans.GetObject(entry.Value, OpenMode.ForRead);
                    var btrId = layout.BlockTableRecordId;
                    if (btrId != ObjectId.Null)
                        return (BlockTableRecord)sourceTrans.GetObject(btrId, OpenMode.ForRead);
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] Error getting Paper Space: {ex.Message}");
            }

            AcadLogger.Log("[LayoutMerger] Fallback to Model Space");
            return (BlockTableRecord)sourceTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(sourceDb), OpenMode.ForRead);
        }

        private List<ViewportInfo> CollectModelViewportInfos(Transaction trans, BlockTableRecord paperSpace, string label)
        {
            var viewports = new List<ViewportInfo>();
            var rawViewports = new List<Viewport>();
            int skippedCount = 0;
            double maxViewArea = 0.0;

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForRead, false) as Viewport;
                    if (vp == null)
                        continue;

                    if (!IsRawModelViewport(vp))
                    {
                        skippedCount++;
                        AcadLogger.LogInfo($"{label}: skipped viewport handle={vp.Handle} paperSize=({vp.Width:F2},{vp.Height:F2}) viewHeight={vp.ViewHeight:F2} scale={vp.CustomScale:F8}");
                        continue;
                    }

                    rawViewports.Add(vp);
                    maxViewArea = Math.Max(maxViewArea, GetViewportViewArea(vp));
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"{label}: failed to scan viewport {id}: {ex.Message}");
                }
            }

            foreach (var vp in rawViewports)
            {
                try
                {
                    if (IsUtilityViewport(vp, maxViewArea))
                    {
                        skippedCount++;
                        AcadLogger.LogInfo($"{label}: skipped utility viewport handle={vp.Handle} paperSize=({vp.Width:F2},{vp.Height:F2}) viewHeight={vp.ViewHeight:F2} scale={vp.CustomScale:F8} viewArea={GetViewportViewArea(vp):F2} maxArea={maxViewArea:F2}");
                        continue;
                    }

                    viewports.Add(new ViewportInfo
                    {
                        SourceId = vp.ObjectId,
                        ViewCenter = vp.ViewCenter,
                        ViewTarget = vp.ViewTarget,
                        ViewDirection = vp.ViewDirection,
                        ViewHeight = vp.ViewHeight,
                        CustomScale = vp.CustomScale,
                        TwistAngle = vp.TwistAngle,
                        CenterPoint = vp.CenterPoint,
                        Width = vp.Width,
                        Height = vp.Height,
                        Number = vp.Number,
                        Locked = vp.Locked,
                        On = vp.On,
                        Layer = vp.Layer,
                        ColorIndex = vp.Color == null ? (short)256 : vp.Color.ColorIndex,
                        LinetypeId = vp.LinetypeId,
                        LineWeight = vp.LineWeight
                    });

                    LogViewportState(label, vp);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"{label}: failed to collect viewport {vp.ObjectId}: {ex.Message}");
                }
            }

            AcadLogger.LogInfo($"{label}: collected={viewports.Count}, skipped={skippedCount}");
            return viewports;
        }

        private void TransformExistingViewportsForPaperRotation(
            Transaction trans,
            BlockTableRecord paperSpace,
            PaperContextResult paperCtx,
            string layoutName)
        {
            if (paperCtx == null)
                return;

            double shiftX = paperCtx.ShiftX;
            double shiftY = paperCtx.ShiftY;

            // Only apply paper normalization shift (no rotation).
            // Paper rotation is disabled to keep title block in correct position.
            // Viewport TwistAngle stays unchanged - content displays upright.
            if (Math.Abs(shiftX) < 1e-6 && Math.Abs(shiftY) < 1e-6)
                return;

            int transformed = 0;
            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForWrite, false) as Viewport;
                    if (vp == null || !IsModelViewportCandidate(vp))
                        continue;

                    double srcCx = vp.CenterPoint.X;
                    double srcCy = vp.CenterPoint.Y;

                    vp.CenterPoint = new Point3d(
                        srcCx + shiftX,
                        srcCy + shiftY,
                        vp.CenterPoint.Z);

                    transformed++;
                    AcadLogger.LogInfo(
                        $"TRANSFORM VP: {layoutName} handle={vp.Handle} " +
                        $"paperShift=({shiftX:F4},{shiftY:F4}) " +
                        $"srcCenter=({srcCx:F2},{srcCy:F2}) => newCenter={FormatPoint(vp.CenterPoint)}");
                }
                catch { }
            }

            if (transformed > 0)
                AcadLogger.LogInfo($"TRANSFORM VP: {layoutName} shifted={transformed} viewport(s) by ({shiftX:F4},{shiftY:F4})");
        }

        private int RecreateLayoutViewports(
            Transaction trans,
            BlockTableRecord paperSpace,
            IReadOnlyList<ViewportInfo> sourceViewports,
            string layoutName,
            Vector3d modelOffset)
        {
            int createdCount = 0;
            if (paperSpace == null || sourceViewports == null || sourceViewports.Count == 0)
                return 0;

            var db = paperSpace.Database;
            var previousWorkingDb = HostApplicationServices.WorkingDatabase;
            bool previousTileMode = db.TileMode;

            try
            {
                HostApplicationServices.WorkingDatabase = db;

                try
                {
                    db.TileMode = false;
                    LayoutManager.Current.CurrentLayout = layoutName;
                }
                catch (System.Exception modeEx)
                {
                    AcadLogger.LogWarning(
                        $"RECREATE viewport could not switch/activate layout '{layoutName}': {modeEx.Message}");
                }

                var viewportLayerId = EnsureViewportLayer(db, trans);

                foreach (var sourceViewport in sourceViewports)
                {
                    try
                    {
                        if (sourceViewport.Width <= 0.0 ||
                            sourceViewport.Height <= 0.0 ||
                            sourceViewport.ViewHeight <= 0.0)
                        {
                            AcadLogger.LogWarning(
                                $"RECREATE skip invalid source viewport '{layoutName}' " +
                                $"paperSize=({sourceViewport.Width:F2},{sourceViewport.Height:F2}) " +
                                $"viewHeight={sourceViewport.ViewHeight:F2}");
                            continue;
                        }

                        var vp = new Viewport();
                        vp.SetDatabaseDefaults(db);
                        paperSpace.AppendEntity(vp);
                        trans.AddNewlyCreatedDBObject(vp, true);

                        vp.LayerId = viewportLayerId;
                        vp.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                            Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);

                        vp.Linetype = "ByLayer";
                        vp.LineWeight = sourceViewport.LineWeight;

                        vp.CenterPoint = sourceViewport.CenterPoint;
                        vp.Width = sourceViewport.Width;
                        vp.Height = sourceViewport.Height;

                        vp.ViewDirection = sourceViewport.ViewDirection.Length == 0.0
                            ? Vector3d.ZAxis
                            : sourceViewport.ViewDirection;

                        // Offset the camera target in WCS with moved ModelSpace geometry.
                        // Keep source ViewCenter so paper viewport framing and CustomScale stay identical.
                        vp.ViewTarget = new Point3d(
                            sourceViewport.ViewTarget.X + modelOffset.X,
                            sourceViewport.ViewTarget.Y + modelOffset.Y,
                            sourceViewport.ViewTarget.Z + modelOffset.Z);
                        vp.ViewCenter = sourceViewport.ViewCenter;
                        vp.ViewHeight = sourceViewport.ViewHeight;

                        if (sourceViewport.CustomScale > 0.0)
                            vp.CustomScale = sourceViewport.CustomScale;

                        vp.TwistAngle = sourceViewport.TwistAngle;

                        vp.PerspectiveOn = false;
                        vp.FrontClipOn = false;
                        vp.BackClipOn = false;
                        vp.NonRectClipOn = false;
                        vp.Locked = false;

                        // Preserve source ON/OFF state instead of forcing ON.
                        vp.On = sourceViewport.On;
                        vp.UpdateDisplay();

                        vp.Locked = sourceViewport.Locked;

                        double halfH = vp.ViewHeight / 2.0;
                        double halfW = halfH * (vp.Width / Math.Max(1e-9, vp.Height));
                        var srcWcsCenter = new Point2d(
                            sourceViewport.ViewTarget.X + sourceViewport.ViewCenter.X,
                            sourceViewport.ViewTarget.Y + sourceViewport.ViewCenter.Y);
                        var newWcsCenter = new Point2d(
                            vp.ViewTarget.X + vp.ViewCenter.X,
                            vp.ViewTarget.Y + vp.ViewCenter.Y);

                        AcadLogger.LogInfo(
                            $"RECREATE viewport: {layoutName} " +
                            $"paperCenter={FormatPoint(vp.CenterPoint)} " +
                            $"paperSize=({vp.Width:F2},{vp.Height:F2}) " +
                            $"srcViewCenter={FormatPoint(sourceViewport.ViewCenter)} " +
                            $"viewCenter={FormatPoint(vp.ViewCenter)} " +
                            $"srcViewTarget={FormatPoint(sourceViewport.ViewTarget)} " +
                            $"viewTarget={FormatPoint(vp.ViewTarget)} " +
                            $"srcWcsCenter={FormatPoint(srcWcsCenter)} " +
                            $"newWcsCenter={FormatPoint(newWcsCenter)} " +
                            $"wcsVisible=({newWcsCenter.X - halfW:F2},{newWcsCenter.Y - halfH:F2}) to " +
                            $"({newWcsCenter.X + halfW:F2},{newWcsCenter.Y + halfH:F2}) " +
                            $"modelOffset={FormatVector(modelOffset)} sourceOn={sourceViewport.On} " +
                            $"newOn={vp.On} vpNumber={vp.Number} locked={vp.Locked}");

                        LogViewportState($"RECREATE after: {layoutName}", vp);
                        createdCount++;
                    }
                    catch (System.Exception ex)
                    {
                        AcadLogger.LogWarning($"RECREATE viewport error: {layoutName}: {ex.Message}");
                    }
                }
                try
                {
                    db.UpdateExt(true);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"RECREATE UpdateExt failed for '{layoutName}': {ex.Message}");
                }
            }
            finally
            {
                try { db.TileMode = previousTileMode; } catch { }
                try { HostApplicationServices.WorkingDatabase = previousWorkingDb; } catch { }
            }

            AcadLogger.LogInfo($"RECREATE summary: {layoutName} created={createdCount}");
            return createdCount;
        }
        private void LogViewportCollection(Transaction trans, BlockTableRecord paperSpace, string label)
        {
            int total = 0;
            int modelViewports = 0;

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForRead, false) as Viewport;
                    if (vp == null)
                        continue;

                    total++;
                    if (IsModelViewportCandidate(vp))
                        modelViewports++;

                    LogViewportState(label, vp);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"{label}: failed to read viewport {id}: {ex.Message}");
                }
            }

            AcadLogger.LogInfo($"{label}: viewport count total={total}, modelViewports={modelViewports}");
        }

        private bool IsModelViewportCandidate(Viewport vp)
        {
            // Keep real model viewports and skip utility/default paper viewport patterns.
            return IsRawModelViewport(vp) && !IsUtilityViewport(vp, 0.0);
        }

        private bool IsRawModelViewport(Viewport vp)
        {
            if (vp == null)
                return false;

            if (vp.Number == 1)
                return false;

            return vp.Width > 0.0 && vp.Height > 0.0 && vp.ViewHeight > 0.0 && vp.CustomScale > 0.0;
        }

        private bool IsUtilityViewport(Viewport vp, double maxViewArea)
        {
            if (!IsRawModelViewport(vp))
                return false;

            bool defaultPaperViewport =
                Math.Abs(vp.Width - 12.0) <= 0.5 &&
                Math.Abs(vp.Height - 9.0) <= 0.5 &&
                Math.Abs(vp.CenterPoint.X - 6.0) <= 0.5 &&
                Math.Abs(vp.CenterPoint.Y - 4.5) <= 0.5 &&
                vp.ViewHeight > 0.0 &&
                vp.ViewHeight <= 25.0 &&
                vp.CustomScale >= 0.5;

            // Guard against false positives: real Revit viewports can be centered near paper center
            // with the same paper size, but their model view center/target are typically not near origin.
            if (!defaultPaperViewport)
                return false;

            // Be more tolerant here because many exports keep the utility viewport center at (6,4.5)
            // while still being the default paper-space helper viewport.
            bool looksLikeDefaultView =
                Math.Abs(vp.ViewCenter.X) <= 10.0 &&
                Math.Abs(vp.ViewCenter.Y) <= 10.0 &&
                Math.Abs(vp.ViewTarget.X) <= 10.0 &&
                Math.Abs(vp.ViewTarget.Y) <= 10.0;

            return looksLikeDefaultView;
        }

        private double GetViewportViewArea(Viewport vp)
        {
            if (vp == null || vp.Width <= 0.0 || vp.Height <= 0.0 || vp.ViewHeight <= 0.0)
                return 0.0;

            return vp.ViewHeight * (vp.ViewHeight * (vp.Width / vp.Height));
        }

        private Vector2d GetViewCenterOffset(Viewport vp, Vector3d modelOffset)
        {
            var directOffset = new Vector2d(modelOffset.X, modelOffset.Y);

            try
            {
                var wcsToDcs = GetWorldToDcsTransform(vp);
                var origin = Point3d.Origin.TransformBy(wcsToDcs);
                var offsetPoint = new Point3d(modelOffset.X, modelOffset.Y, modelOffset.Z).TransformBy(wcsToDcs);
                var delta = origin.GetVectorTo(offsetPoint);

                var transformedOffset = new Vector2d(delta.X, delta.Y);
                double directLen = directOffset.Length;
                double transformedLen = transformedOffset.Length;

                if (double.IsNaN(transformedLen) || double.IsInfinity(transformedLen) ||
                    (directLen > 1e-9 && transformedLen > directLen * 100.0))
                {
                    AcadLogger.LogWarning(
                        $"GetViewCenterOffset: unstable transformed delta " +
                        $"({transformedOffset.X:F4},{transformedOffset.Y:F4}), " +
                        $"fallback direct ({directOffset.X:F4},{directOffset.Y:F4})");
                    return directOffset;
                }

                AcadLogger.LogDebug(
                    $"GetViewCenterOffset: transformed=({transformedOffset.X:F4},{transformedOffset.Y:F4}), " +
                    $"direct=({directOffset.X:F4},{directOffset.Y:F4})");

                return transformedOffset;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"GetViewCenterOffset fallback: {ex.Message}");
                return directOffset;
            }
        }
        private Matrix3d GetWorldToDcsTransform(Viewport vp)
        {
            var viewDirection = vp.ViewDirection;
            if (viewDirection.Length == 0.0)
                viewDirection = Vector3d.ZAxis;

            var target = vp.ViewTarget;
            var transform = Matrix3d.PlaneToWorld(viewDirection);
            transform = Matrix3d.Displacement(target - Point3d.Origin) * transform;
            transform = Matrix3d.Rotation(-vp.TwistAngle, viewDirection, target) * transform;
            return transform.Inverse();
        }

        private void LogViewportState(string label, Viewport vp)
        {
            double viewWidth = 0.0;
            if (vp.Height > 0)
                viewWidth = vp.ViewHeight * (vp.Width / vp.Height);

            AcadLogger.LogInfo(
                $"{label} VP#{vp.Number} handle={vp.Handle} " +
                $"paperCenter={FormatPoint(vp.CenterPoint)} paperSize=({vp.Width:F2},{vp.Height:F2}) " +
                $"viewCenter={FormatPoint(vp.ViewCenter)} viewTarget={FormatPoint(vp.ViewTarget)} " +
                $"viewDir={FormatVector(vp.ViewDirection)} viewSize=({viewWidth:F2},{vp.ViewHeight:F2}) " +
                $"customScale={vp.CustomScale:F8} twist={vp.TwistAngle:F8} locked={vp.Locked} on={vp.On} " +
                $"visible={FormatExtents(GetViewportViewExtents(vp.ViewCenter, vp))}");
        }

        private void LogLayoutDiag(
            string layoutName,
            string cloneMode,
            int sourceUsableViewports,
            int fixedViewports,
            int bakedEntities,
            int erasedViewports,
            Vector3d modelOffset,
            bool keepViewportLive)
        {
            AcadLogger.LogInfo(
                $"[LAYOUT-DIAG] {{\"layout\":\"{layoutName}\",\"cloneMode\":\"{cloneMode}\",\"keepViewportLive\":{keepViewportLive.ToString().ToLowerInvariant()},\"srcUsableVp\":{sourceUsableViewports},\"fixedVp\":{fixedViewports},\"bakedEntities\":{bakedEntities},\"erasedVp\":{erasedViewports},\"msOffset\":\"{FormatVector(modelOffset)}\"}}");
        }

        private string FormatPoint(Point2d point)
        {
            return $"({point.X:F4},{point.Y:F4})";
        }

        private string FormatPoint(Point3d point)
        {
            return $"({point.X:F4},{point.Y:F4},{point.Z:F4})";
        }

        private string FormatVector(Vector3d vector)
        {
            return $"({vector.X:F4},{vector.Y:F4},{vector.Z:F4})";
        }

        private Extents3d GetViewportViewExtents(Point2d center, Viewport vp)
        {
            if (vp == null || vp.Width <= 0.0 || vp.Height <= 0.0 || vp.ViewHeight <= 0.0)
                return new Extents3d(new Point3d(center.X, center.Y, 0), new Point3d(center.X, center.Y, 0));

            double viewWidth = vp.ViewHeight * (vp.Width / vp.Height);
            double halfWidth = viewWidth / 2.0;
            double halfHeight = vp.ViewHeight / 2.0;

            return new Extents3d(
                new Point3d(center.X - halfWidth, center.Y - halfHeight, 0),
                new Point3d(center.X + halfWidth, center.Y + halfHeight, 0));
        }

        private string FormatExtents(Extents3d extents)
        {
            return $"min={FormatPoint(extents.MinPoint)}, max={FormatPoint(extents.MaxPoint)}, " +
                $"size=({extents.MaxPoint.X - extents.MinPoint.X:F4},{extents.MaxPoint.Y - extents.MinPoint.Y:F4})";
        }

        private Extents3d GetLayoutModelViewExtents(Database db, Transaction trans, Layout layout, Extents3d fallback, string label)
        {
            if (layout == null || layout.BlockTableRecordId == ObjectId.Null)
                return fallback;

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            int viewportCount = 0;

            try
            {
                var paperSpace = (BlockTableRecord)trans.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                var rawViewports = new List<Viewport>();
                double maxViewArea = 0.0;

                foreach (ObjectId id in paperSpace)
                {
                    var vp = trans.GetObject(id, OpenMode.ForRead, false) as Viewport;
                    if (!IsRawModelViewport(vp))
                        continue;

                    rawViewports.Add(vp);
                    maxViewArea = Math.Max(maxViewArea, GetViewportViewArea(vp));
                }

                foreach (var vp in rawViewports)
                {
                    if (IsUtilityViewport(vp, maxViewArea))
                    {
                        AcadLogger.LogDebug($"{label}: ignored utility viewport handle={vp.Handle} paperSize=({vp.Width:F2},{vp.Height:F2}) viewArea={GetViewportViewArea(vp):F2} maxArea={maxViewArea:F2}");
                        continue;
                    }

                    var vpExtents = GetViewportViewExtents(vp.ViewCenter, vp);

                    minX = Math.Min(minX, vpExtents.MinPoint.X);
                    maxX = Math.Max(maxX, vpExtents.MaxPoint.X);
                    minY = Math.Min(minY, vpExtents.MinPoint.Y);
                    maxY = Math.Max(maxY, vpExtents.MaxPoint.Y);
                    viewportCount++;
                    AcadLogger.LogDebug($"{label}: viewport source window {FormatExtents(vpExtents)} handle={vp.Handle}");
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"GetLayoutModelViewExtents {label}: {ex.Message}");
            }

            if (viewportCount == 0 || minX == double.MaxValue)
            {
                AcadLogger.LogWarning($"{label}: no model viewport extents found; using ModelSpace extents for spacing");
                return fallback;
            }

            var extents = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
            double width = maxX - minX;
            double height = maxY - minY;
            AcadLogger.LogInfo($"{label}: viewport view extents count={viewportCount}, width={width:F2}, height={height:F2}, minX={minX:F2}, maxX={maxX:F2}");
            return extents;
        }

        private Extents3d CombineExtents(Extents3d a, Extents3d b)
        {
            double minX = Math.Min(a.MinPoint.X, b.MinPoint.X);
            double minY = Math.Min(a.MinPoint.Y, b.MinPoint.Y);
            double maxX = Math.Max(a.MaxPoint.X, b.MaxPoint.X);
            double maxY = Math.Max(a.MaxPoint.Y, b.MaxPoint.Y);
            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        }

        private double GetLayoutGap(Extents3d extents)
        {
            double width = extents.MaxPoint.X - extents.MinPoint.X;
            if (width <= 0)
                width = 100000;

            return Math.Max(LayoutSpacing * 10, width * 0.5);
        }

        private DwgGeometryStats GetModelSpaceStats(Database db, Transaction trans)
        {
            string filePath = db.Filename ?? string.Empty;
            string cacheKey = GetExtentsCacheKey(filePath, "Model");

            if (TryGetExtentsCache(cacheKey, out var cached))
            {
                return new DwgGeometryStats
                {
                    EntityCount = cached.EntityCount,
                    ExtentsEntityCount = cached.ExtentsEntityCount,
                    Extents = cached.Extents
                };
            }

            var modelSpace = (BlockTableRecord)trans.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

            int entityCount = 0;
            foreach (ObjectId id in modelSpace)
            {
                if (!id.IsNull && !id.IsErased)
                    entityCount++;
            }

            int extentsEntityCount;
            var extents = GetExtents(modelSpace, trans, out extentsEntityCount);

            SetExtentsCache(cacheKey, new ExtentsCacheEntry
            {
                Extents = extents,
                EntityCount = entityCount,
                ExtentsEntityCount = extentsEntityCount,
                CachedAt = DateTime.Now,
                IsModelSpace = true
            });

            return new DwgGeometryStats
            {
                EntityCount = entityCount,
                ExtentsEntityCount = extentsEntityCount,
                Extents = extents
            };
        }

        private void LogModelSpaceStats(string label, DwgGeometryStats stats)
        {
            double width = stats.Extents.MaxPoint.X - stats.Extents.MinPoint.X;
            double height = stats.Extents.MaxPoint.Y - stats.Extents.MinPoint.Y;
            AcadLogger.LogInfo(
                $"{label}: MS entities={stats.EntityCount}, extentsEntities={stats.ExtentsEntityCount}, " +
                $"width={width:F2}, height={height:F2}, " +
                $"min=({stats.Extents.MinPoint.X:F2},{stats.Extents.MinPoint.Y:F2}), " +
                $"max=({stats.Extents.MaxPoint.X:F2},{stats.Extents.MaxPoint.Y:F2})");

            if (stats.EntityCount == 0 || stats.ExtentsEntityCount == 0 || (Math.Abs(width) < 1e-6 && Math.Abs(height) < 1e-6))
            {
                AcadLogger.LogWarning($"{label}: ModelSpace still looks empty after XREF bind; layout may not display sheet geometry correctly");
            }
        }

        private List<LayoutVerifyStats> InspectPaperLayouts(Database db, Transaction tr)
        {
            var result = new List<LayoutVerifyStats>();
            var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            foreach (DBDictionaryEntry entry in layouts)
            {
                try
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (layout == null || layout.ModelType || layout.BlockTableRecordId.IsNull)
                        continue;

                    var paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    var stats = new LayoutVerifyStats { Name = layout.LayoutName };

                    foreach (ObjectId id in paperSpace)
                    {
                        try
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                            if (ent == null || ent.IsErased)
                                continue;

                            stats.EntityCount++;

                            if (string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                            {
                                stats.BackgroundEntityCount++;
                                continue;
                            }

                            if (ent is Viewport)
                            {
                                stats.ViewportEntityCount++;
                                continue;
                            }

                            stats.ContentEntityCount++;
                        }
                        catch
                        {
                        }
                    }

                    result.Add(stats);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"VERIFY: failed to inspect layout '{entry.Key}': {ex.Message}");
                }
            }

            return result;
        }

        private bool IsDefaultLayoutName(string layoutName)
        {
            return string.Equals(layoutName, "Layout1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(layoutName, "Layout2", StringComparison.OrdinalIgnoreCase);
        }

        private bool LayoutHasContent(BlockTableRecord paperSpace, Transaction tr)
        {
            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (entity == null || entity.IsErased)
                        continue;

                    if (entity is Viewport)
                        continue;

                    if (string.Equals(entity.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private void ApplyLayoutPlotSettingsSafely(
            Layout destLayout,
            PlotSettings sourcePlotSettings,
            string layoutName,
            ObjectId originalBtrId,
            int originalTabOrder,
            string phaseTag)
        {
            if (destLayout == null)
                return;

            try
            {
                if (sourcePlotSettings != null)
                {
                    // Log source paper size before copy so we can detect silent failures
                    AcadLogger.LogInfo(
                        $"{phaseTag}: Source paper size=({sourcePlotSettings.PlotPaperSize.X:F2},{sourcePlotSettings.PlotPaperSize.Y:F2}) for '{layoutName}'");

                    destLayout.CopyFrom(sourcePlotSettings);

                    double copyW = destLayout.PlotPaperSize.X;
                    double copyH = destLayout.PlotPaperSize.Y;
                    AcadLogger.LogInfo(
                        $"{phaseTag}: Dest paper size=({copyW:F2},{copyH:F2}) for '{layoutName}'");

                    try
                    {
                        EnsurePlotSettingsRefreshed(destLayout);
                        var psv = PlotSettingsValidator.Current;

                        if (!string.IsNullOrWhiteSpace(sourcePlotSettings.PlotConfigurationName) &&
                            !string.IsNullOrWhiteSpace(sourcePlotSettings.CanonicalMediaName))
                        {
                            psv.SetPlotConfigurationName(
                                destLayout,
                                sourcePlotSettings.PlotConfigurationName,
                                sourcePlotSettings.CanonicalMediaName);
                        }

                        try { psv.SetPlotType(destLayout, sourcePlotSettings.PlotType); } catch { }
                        try { psv.SetPlotOrigin(destLayout, sourcePlotSettings.PlotOrigin); } catch { }
                        try { psv.SetPlotRotation(destLayout, sourcePlotSettings.PlotRotation); } catch { }
                        try { psv.SetPlotPaperUnits(destLayout, sourcePlotSettings.PlotPaperUnits); } catch { }

                        copyW = destLayout.PlotPaperSize.X;
                        copyH = destLayout.PlotPaperSize.Y;
                        AcadLogger.LogInfo(
                            $"{phaseTag}: PSV normalized page setup for '{layoutName}', " +
                            $"device='{destLayout.PlotConfigurationName}', media='{destLayout.CanonicalMediaName}', " +
                            $"paper=({copyW:F2},{copyH:F2})");
                    }
                    catch (System.Exception psvNormalizeEx)
                    {
                        AcadLogger.LogWarning(
                            $"{phaseTag}: PSV page setup normalize failed for '{layoutName}': {psvNormalizeEx.Message}");
                    }

                    // If CopyFrom silently produced (0,0), apply the canonical paper size
                    // via PlotSettingsValidator so AutoCAD fully registers it.
                    if (copyW < 1.0 || copyH < 1.0)
                    {
                        double srcW = sourcePlotSettings.PlotPaperSize.X;
                        double srcH = sourcePlotSettings.PlotPaperSize.Y;
                        if (srcW > 1.0 && srcH > 1.0)
                        {
                            try
                            {
                                var psv = PlotSettingsValidator.Current;
                                psv.SetPlotConfigurationName(destLayout,
                                    sourcePlotSettings.PlotConfigurationName,
                                    sourcePlotSettings.CanonicalMediaName);
                                AcadLogger.LogInfo(
                                    $"{phaseTag}: PSV applied canonical media '{sourcePlotSettings.CanonicalMediaName}' " +
                                    $"for '{layoutName}', result size=({destLayout.PlotPaperSize.X:F2},{destLayout.PlotPaperSize.Y:F2})");
                            }
                            catch (System.Exception psvEx)
                            {
                                AcadLogger.LogWarning($"{phaseTag}: PSV fallback failed for '{layoutName}': {psvEx.Message}");
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Keep merge alive when source layout/plot settings are corrupt.
                AcadLogger.LogWarning($"{phaseTag}: CopyFrom PlotSettings failed for '{layoutName}': {ex.Message}");
            }

            try
            {
                if (!originalBtrId.IsNull)
                    destLayout.BlockTableRecordId = originalBtrId;
                if (!string.IsNullOrWhiteSpace(layoutName))
                    destLayout.LayoutName = layoutName;
                destLayout.TabOrder = originalTabOrder;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"{phaseTag}: Restore layout identity failed for '{layoutName}': {ex.Message}");
            }
        }

        private void DisposePlotSettings(IEnumerable<SourceFileInfo> infos)
        {
            if (infos == null)
                return;

            foreach (var info in infos)
            {
                if (info?.PlotSettings == null)
                    continue;

                try
                {
                    info.PlotSettings.Dispose();
                }
                catch
                {
                }
                finally
                {
                    info.PlotSettings = null;
                }
            }
        }

        private int CountModelSpaceBackgrounds(Database db, Transaction tr)
        {
            int count = 0;
            var modelSpace = (BlockTableRecord)tr.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent != null &&
                        !ent.IsErased &&
                        string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        count++;
                    }
                }
                catch
                {
                }
            }

            // Each sheet background is currently created as a boundary polyline plus a solid hatch.
            return (int)Math.Ceiling(count / 2.0);
        }

        private List<RasterImageInfo> ScanRasterImages(string dwgPath)
        {
            var result = new List<RasterImageInfo>();
            var db = new Database(false, true);

            using (db)
            {
                db.ReadDwgFile(dwgPath, FileShare.ReadWrite, true, "");
                db.CloseInput(true);

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = tr.GetObject(btrId, OpenMode.ForRead, false) as BlockTableRecord;
                        if (btr == null)
                            continue;

                        foreach (ObjectId id in btr)
                        {
                            try
                            {
                                var raster = tr.GetObject(id, OpenMode.ForRead, false) as RasterImage;
                                if (raster == null || raster.IsErased)
                                    continue;

                                result.Add(new RasterImageInfo
                                {
                                    Handle = raster.Handle.ToString(),
                                    Layer = raster.Layer,
                                    Owner = btr.Name
                                });
                            }
                            catch
                            {
                            }
                        }
                    }

                    tr.Commit();
                }
            }

            return result;
        }

        private Extents3d GetExtents(BlockTableRecord btr, string cacheKey = null)
        {
            if (!string.IsNullOrEmpty(cacheKey) && TryGetExtentsCache(cacheKey, out var cached))
            {
                return cached.Extents;
            }

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (ObjectId id in btr)
            {
                try
                {
                    var ent = (Entity)id.GetObject(OpenMode.ForRead);
                    if (ent != null)
                    {
                        var ext = ent.GeometricExtents;
                        if (ext.MinPoint.X < minX) minX = ext.MinPoint.X;
                        if (ext.MinPoint.Y < minY) minY = ext.MinPoint.Y;
                        if (ext.MaxPoint.X > maxX) maxX = ext.MaxPoint.X;
                        if (ext.MaxPoint.Y > maxY) maxY = ext.MaxPoint.Y;
                    }
                }
                catch { }
            }

            Extents3d result;
            if (minX == double.MaxValue)
                result = new Extents3d(new Point3d(0, 0, 0), new Point3d(0, 0, 0));
            else
                result = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));

            if (!string.IsNullOrEmpty(cacheKey))
            {
                SetExtentsCache(cacheKey, new ExtentsCacheEntry
                {
                    Extents = result,
                    EntityCount = 0,
                    ExtentsEntityCount = 0,
                    CachedAt = DateTime.Now,
                    IsModelSpace = false
                });
            }

            return result;
        }

        private Extents3d GetExtents(BlockTableRecord btr, Transaction trans, out int extentsEntityCount, string cacheKey = null)
        {
            if (!string.IsNullOrEmpty(cacheKey) && TryGetExtentsCache(cacheKey, out var cached))
            {
                extentsEntityCount = cached.ExtentsEntityCount;
                return cached.Extents;
            }

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            extentsEntityCount = 0;

            foreach (ObjectId id in btr)
            {
                try
                {
                    var ent = trans.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null)
                        continue;

                    var ext = ent.GeometricExtents;
                    if (ext.MinPoint.X < minX) minX = ext.MinPoint.X;
                    if (ext.MinPoint.Y < minY) minY = ext.MinPoint.Y;
                    if (ext.MaxPoint.X > maxX) maxX = ext.MaxPoint.X;
                    if (ext.MaxPoint.Y > maxY) maxY = ext.MaxPoint.Y;
                    extentsEntityCount++;
                }
                catch
                {
                }
            }

            Extents3d result;
            if (minX == double.MaxValue)
                result = new Extents3d(new Point3d(0, 0, 0), new Point3d(0, 0, 0));
            else
                result = new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));

            if (!string.IsNullOrEmpty(cacheKey))
            {
                SetExtentsCache(cacheKey, new ExtentsCacheEntry
                {
                    Extents = result,
                    EntityCount = extentsEntityCount,
                    ExtentsEntityCount = extentsEntityCount,
                    CachedAt = DateTime.Now,
                    IsModelSpace = false
                });
            }

            return result;
        }

        private DwgVersion GetDwgVersion(string version)
        {
            if (string.IsNullOrEmpty(version) || version == "Current")
                return DwgVersion.Current;

            switch (version)
            {
                case "2018": return DwgVersion.AC1032;
                case "2013": return DwgVersion.AC1027;
                case "2010": return DwgVersion.AC1024;
                case "2007": return DwgVersion.AC1021;
                case "AC1027": return DwgVersion.AC1027;
                case "AC1024": return DwgVersion.AC1024;
                case "AC1021": return DwgVersion.AC1021;
                case "AC1015": return DwgVersion.AC1015;
                case "AC1014": return DwgVersion.AC1014;
                case "AC1012": return DwgVersion.AC1012;
                default:
                    AcadLogger.Log($"[LayoutMerger] Unknown DwgVersion '{version}', using Current");
                    return DwgVersion.Current;
            }
        }

        private class SourceFileInfo
    {
        public string FilePath;
        public string LayoutName;
        public Vector3d MsOffset;
        public Extents3d MsExtents;
        public bool ModelType;
        public PlotSettings PlotSettings;
    }

        private class DwgGeometryStats
        {
            public int EntityCount;
            public int ExtentsEntityCount;
            public Extents3d Extents;
        }

        private class LayoutVerifyStats
        {
            public string Name;
            public int EntityCount;
            public int ContentEntityCount;
            public int BackgroundEntityCount;
            public int ViewportEntityCount;
        }

        private class LayoutRegenInfo
        {
            public string Name;
            public int TabOrder;
            public int PaperEntityCount;
            public int ViewportCount;
            public int ExtentsEntityCount;
            public Extents3d PaperExtents;
            public bool RequiresRegen;
        }

        private class RasterImageInfo
        {
            public string Handle;
            public string Layer;
            public string Owner;
        }

    private class PaperContextResult
        {
            public PlotRotation AppliedRotation { get; set; }
            public double RequiredWidth { get; set; }
            public double RequiredHeight { get; set; }
            public bool WasAdjusted { get; set; }
            public double ShiftX { get; set; }
            public double ShiftY { get; set; }
        }

    private class ViewportInfo
        {
            public ObjectId SourceId { get; set; }
            public Point2d ViewCenter { get; set; }
            public Point3d ViewTarget { get; set; }
            public Vector3d ViewDirection { get; set; }
            public double ViewHeight { get; set; }
            public double CustomScale { get; set; }
            public double TwistAngle { get; set; }
            public Point3d CenterPoint { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public int Number { get; set; }
            public bool Locked { get; set; }
            public bool On { get; set; }
            public string Layer { get; set; }
            public short ColorIndex { get; set; }
            public ObjectId LinetypeId { get; set; }
            public LineWeight LineWeight { get; set; }
        }
    }
}




