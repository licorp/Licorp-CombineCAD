using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        private const double PaperBackgroundFallbackWidth = 1066.8;
        private const double PaperBackgroundFallbackHeight = 762.0;
        private const double ModelSpaceSheetMinGap = 100.0;
        private const int AutoCadLayoutNameMaxLength = 31;
        
        // Track ModelSpace offset for each source file
        private Dictionary<string, Vector3d> _msOffsets = new Dictionary<string, Vector3d>();
        private double _currentMsXOffset = 0.0;

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

                // Tìm base file
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
                var sourceInfos = new List<SourceFileInfo>();
                var pendingScheduleOnlyLayouts = new List<SourceFileInfo>();
                var usedLayoutNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                using (outputDb)
                {
                    outputDb.ReadDwgFile(baseFile, FileShare.ReadWrite, true, "");
                    outputDb.CloseInput(true);
                    AcadLogger.LogInfo("Base file opened successfully");
                    BindXrefsSafe(outputDb);

                    // === PHẦN KHỞI TẠO: Xử lý file đầu tiên ===
                    using (var trans = outputDb.TransactionManager.StartTransaction())
                    {
                        var firstSource = config.SourceFiles.First(s => File.Exists(s.Path));
                        var firstLayoutName = GetSafeAutoCadLayoutName(firstSource.Layout ?? "Layout1", usedLayoutNames);
                        
                        // Rename base layout
                        if (!string.IsNullOrEmpty(firstSource.Layout))
                        {
                            RenameLayoutInDb(outputDb, "Layout1", firstLayoutName);
                            AcadLogger.LogInfo($"Renamed base layout to '{firstLayoutName}'");
                        }

                        // Tính extents của base ModelSpace → khởi tạo offset
                        var baseMsId = SymbolUtilityServices.GetBlockModelSpaceId(outputDb);
                        var baseMs = (BlockTableRecord)trans.GetObject(baseMsId, OpenMode.ForRead);
                        var baseStats = GetModelSpaceStats(outputDb, trans);
                        LogModelSpaceStats("Base after bind", baseStats);
                        var baseExtents = baseStats.Extents;
                        double baseWidth = baseExtents.MaxPoint.X - baseExtents.MinPoint.X;

                        // Validate: đảm bảo width > 0
                        if (baseWidth <= 0)
                        {
                            AcadLogger.LogWarning($"Base MS extents invalid (width={baseWidth}), using safe default");
                            baseWidth = 100000;
                        }

                        AcadLogger.LogInfo($"Base MS extents width: {baseWidth:F2}");

                        // Lưu thông tin file đầu tiên
                        var firstLayout = GetSourceLayout(outputDb, trans, firstLayoutName);
                        var baseOccupiedExtents = baseExtents;
                        if (firstLayout != null)
                        {
                            baseOccupiedExtents = CombineExtents(
                                baseExtents,
                                GetLayoutModelViewExtents(outputDb, trans, firstLayout, baseExtents, "Base layout"));
                            var firstBtr = (BlockTableRecord)trans.GetObject(firstLayout.BlockTableRecordId, OpenMode.ForWrite);
                            var firstViewports = CollectModelViewportInfos(trans, firstBtr, $"BASE usable viewports: {firstLayoutName}");
                            int baseBakedCount = BakeModelViewsToPaperSpace(
                                outputDb, trans, outputDb, trans, firstBtr, firstViewports, firstLayoutName);
                            int baseErasedViewportCount = baseBakedCount > 0
                                ? EraseAllLayoutViewports(trans, firstBtr, firstLayoutName)
                                : 0;
                            if (baseBakedCount == 0)
                                AcadLogger.LogWarning($"BASE: Paper bake produced no entities; keeping original viewport(s) for '{firstLayoutName}'");
                            AcadLogger.LogInfo(
                                $"BASE: Baked {baseBakedCount} model entity clone(s) to PaperSpace and erased " +
                                $"{baseErasedViewportCount} viewport(s) for '{firstLayoutName}'");
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

                    int clonedCount = 1; // Base layout đã rename
                    int fileIndex = 2;

                    // === VÒNG LẶP: Xử lý file 2 → file N ===
                    foreach (var source in config.SourceFiles)
                    {
                        if (!File.Exists(source.Path))
                        {
                            AcadLogger.LogWarning($"Source not found: {source.Path}");
                            continue;
                        }

                        // Bỏ qua base file
                        if (source.Path.Equals(baseFile, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string requestedName = source.Layout ?? $"Layout{fileIndex}";
                        string desiredName = GetSafeAutoCadLayoutName(requestedName, usedLayoutNames);
                        AcadLogger.LogInfo($"Processing: {Path.GetFileName(source.Path)} → Layout '{desiredName}'");

                        PlotSettings savedPlotSettings = null;
                        bool modelType = false;
                        Vector3d msOffset = Vector3d.ZAxis; // placeholder
                        Extents3d srcExtents = new Extents3d();

                        var sourceDb = new Database(false, true);
                        using (sourceDb)
                        {
                            sourceDb.ReadDwgFile(source.Path, FileShare.ReadWrite, true, "");
                            sourceDb.CloseInput(true);
                            BindXrefsSafe(sourceDb);

                            using (var srcTrans = sourceDb.TransactionManager.StartTransaction())
                            using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                            {
                                // === GIAI ĐOẠN 1: Clone ModelSpace ===
                                // Rename blocks trước khi clone
                                RenameBlocksInDb(sourceDb, srcTrans, $"File{fileIndex}_");

                                // Clone ModelSpace entities
                                var srcMs = (BlockTableRecord)srcTrans.GetObject(
                                    SymbolUtilityServices.GetBlockModelSpaceId(sourceDb), OpenMode.ForRead);
                                var outputMs = (BlockTableRecord)outputTrans.GetObject(
                                    SymbolUtilityServices.GetBlockModelSpaceId(outputDb), OpenMode.ForWrite);
                                var srcStats = GetModelSpaceStats(sourceDb, srcTrans);
                                LogModelSpaceStats($"Source after bind: {Path.GetFileName(source.Path)}", srcStats);

                                Layout srcLayout;
                                srcLayout = GetSourceLayout(sourceDb, srcTrans, source.Layout);
                                if (srcLayout == null)
                                {
                                    AcadLogger.LogError($"No layout found in {Path.GetFileName(source.Path)}");
                                    fileIndex++;
                                    continue;
                                }

modelType = srcLayout.ModelType;
            savedPlotSettings = new PlotSettings(srcLayout.ModelType);
            savedPlotSettings.CopyFrom(srcLayout);

            // FIX Bug 1: Use viewport view extents ONLY for offset calculation
            // Don't combine with ModelSpace extents (which may include garbage entities at far coordinates)
            var viewportViewExtents = GetLayoutModelViewExtents(sourceDb, srcTrans, srcLayout, srcStats.Extents, desiredName);
            var sourceVisibleExtents = CombineExtents(
                srcStats.Extents,
                viewportViewExtents);
            AcadLogger.LogInfo($"GD1: {desiredName} viewport view extents {FormatExtents(viewportViewExtents)}");
            AcadLogger.LogInfo($"GD1: {desiredName} source visible extents (combined) {FormatExtents(sourceVisibleExtents)}");

            var msIds = new ObjectIdCollection();
                                foreach (ObjectId id in srcMs) msIds.Add(id);

                                msOffset = new Vector3d(_currentMsXOffset - viewportViewExtents.MinPoint.X, 0, 0);
                                _msOffsets[desiredName] = msOffset;
                                AcadLogger.LogInfo($"GD1: {desiredName} msOffset={FormatVector(msOffset)}, target visible min X={_currentMsXOffset:F2}");

                                var msIdMap = new IdMapping();
                                int msClonedCount = 0;
                                int msMovedCount = 0;

                                if (msIds.Count > 0)
                                {
                                    // Dùng Replace thay vì Ignore (vì đã rename blocks)
                                    sourceDb.WblockCloneObjects(msIds, outputMs.ObjectId, msIdMap, 
                                        DuplicateRecordCloning.Replace, false);

                                    // Dịch chuyển entities
                                    foreach (ObjectId srcId in msIds)
                                    {
                                        if (msIdMap.Contains(srcId))
                                        {
                                            ObjectId destId = msIdMap[srcId].Value;
                                            if (!destId.IsNull)
                                            {
                                                try
                                                {
                                                    var ent = outputTrans.GetObject(destId, OpenMode.ForWrite) as Entity;
                                                    if (ent != null)
                                                    {
                                                        ent.TransformBy(Matrix3d.Displacement(msOffset));
                                                        msMovedCount++;
                                                    }
                                                    msClonedCount++;
                                                }
                                                catch { }
                                            }
                                        }
                                    }

                                    AcadLogger.LogInfo($"GD1: Cloned {msClonedCount} MS entities, moved {msMovedCount} by offset ({msOffset.X:F2}, 0, 0)");
                                }

                                // Tính extents và cập nhật offset
                                srcExtents = srcStats.Extents;
                                _currentMsXOffset = msOffset.X + sourceVisibleExtents.MaxPoint.X + GetLayoutGap(sourceVisibleExtents);
                                AcadLogger.LogInfo($"GD1: Updated next visible min X to {_currentMsXOffset:F2}");

                                // === GIAI DOAN 2: Clone source Layout object directly ===
                                // This keeps AutoCAD's internal Layout/PaperSpace/Viewport wiring intact.
                                var shouldUseDirectLayoutClone = UseDirectLayoutClone();

                                if (shouldUseDirectLayoutClone)
                                {
                                    var srcBtrIdDirect = srcLayout.BlockTableRecordId;
                                    var srcBtrDirect = (BlockTableRecord)srcTrans.GetObject(srcBtrIdDirect, OpenMode.ForRead);
                                    LogViewportCollection(srcTrans, srcBtrDirect, $"SRC before layout clone: {desiredName}");
                                    var sourceViewportsDirect = CollectModelViewportInfos(srcTrans, srcBtrDirect, $"SRC usable viewports: {desiredName}");

                                    var destBtrIdDirect = CloneLayoutFromSource(sourceDb, outputDb, srcTrans, outputTrans, srcLayout, desiredName);
                                    if (destBtrIdDirect.IsNull)
                                    {
                                        AcadLogger.LogWarning(
                                            $"GD2: Direct layout clone failed for '{desiredName}'. " +
                                            "Falling back to manual layout creation.");
                                    }
                                    else
                                    {
                                        var destBtrDirect = (BlockTableRecord)outputTrans.GetObject(destBtrIdDirect, OpenMode.ForWrite);
                                        var destLayoutDirect = (Layout)outputTrans.GetObject(destBtrDirect.LayoutId, OpenMode.ForWrite);
                                        AcadLogger.LogInfo(
                                            $"GD2: Cloned layout object for '{desiredName}', source usable viewport count={sourceViewportsDirect.Count}, " +
                                            $"paperSize=({destLayoutDirect.PlotPaperSize.X:F2},{destLayoutDirect.PlotPaperSize.Y:F2})");

                                        LogViewportCollection(outputTrans, destBtrDirect, $"DEST after layout clone before fix: {desiredName}");

                                        int bakedCountDirect = BakeModelViewsToPaperSpace(
                                            sourceDb, srcTrans, outputDb, outputTrans, destBtrDirect, sourceViewportsDirect, desiredName);
                                        bool scheduleHasPaperContent =
                                            srcStats.EntityCount == 0 &&
                                            LayoutHasContent(destBtrDirect, outputTrans);
                                        int erasedViewportCountDirect = bakedCountDirect > 0 || scheduleHasPaperContent
                                            ? EraseAllLayoutViewports(outputTrans, destBtrDirect, desiredName)
                                            : 0;
                                        if (bakedCountDirect == 0)
                                        {
                                            if (scheduleHasPaperContent)
                                            {
                                                AcadLogger.LogInfo(
                                                    $"GD2: '{desiredName}' is schedule-only with PaperSpace content; removed viewport(s) so the layout behaves like a normal sheet tab.");
                                            }
                                            else
                                            {
                                                AcadLogger.LogWarning($"GD2: Paper bake produced no entities; keeping original viewport(s) for '{desiredName}'");
                                            }
                                        }
                                        AcadLogger.LogInfo(
                                            $"GD2: Baked {bakedCountDirect} model entity clone(s) to PaperSpace and erased " +
                                            $"{erasedViewportCountDirect} viewport(s) for '{desiredName}'");

                                        if (srcStats.EntityCount == 0 && !LayoutHasContent(destBtrDirect, outputTrans))
                                        {
                                            AddSchedulePlaceholderContent(destBtrDirect, outputTrans, desiredName, destLayoutDirect);
                                        }

                                        LogViewportCollection(outputTrans, destBtrDirect, $"DEST after paper bake: {desiredName}");

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

                                        clonedCount++;
                                        AcadLogger.LogInfo($"Successfully cloned layout '{desiredName}'");
                                        fileIndex++;
                                        continue;
                                    }
                                }

                                // === GIAI ĐOẠN 2: Tạo Layout + Clone PaperSpace ===
                                // Tạo layout mới
var destBtrId = CreateNewLayoutInDb(outputDb, outputTrans, desiredName);
            if (destBtrId.IsNull)
            {
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
            }

            var destBtr = (BlockTableRecord)outputTrans.GetObject(destBtrId, OpenMode.ForWrite);
            var destLayout = (Layout)outputTrans.GetObject(destBtr.LayoutId, OpenMode.ForWrite);

            // FIX: Copy PlotSettings BEFORE cloning PaperSpace entities
            // This ensures the layout has correct paper size before viewport positions are calculated
            var savedBtrId = destLayout.BlockTableRecordId;
            var savedTabOrder = destLayout.TabOrder;
            AcadLogger.LogInfo($"GD2: Source PlotOrigin=({savedPlotSettings.PlotOrigin.X:F2},{savedPlotSettings.PlotOrigin.Y:F2})");
            destLayout.CopyFrom(savedPlotSettings);
            AcadLogger.LogInfo($"GD2: Dest paper size=({destLayout.PlotPaperSize.X:F2},{destLayout.PlotPaperSize.Y:F2})");
            destLayout.BlockTableRecordId = savedBtrId;
            destLayout.LayoutName = desiredName;
            destLayout.TabOrder = savedTabOrder;
            AcadLogger.LogInfo($"GD2: Copied plot settings for '{desiredName}'");

            // Clone PaperSpace entities (now with correct paper size)
            var srcBtrId = srcLayout.BlockTableRecordId;
                                var srcBtr = (BlockTableRecord)srcTrans.GetObject(srcBtrId, OpenMode.ForRead);
                                LogViewportCollection(srcTrans, srcBtr, $"SRC before PS clone: {desiredName}");
                                var sourceViewports = CollectModelViewportInfos(srcTrans, srcBtr, $"SRC usable viewports: {desiredName}");
                                
                                var psIds = new ObjectIdCollection();
                                int psViewportSkippedCount = 0;
                                int psViewportClonedCount = 0;
                                foreach (ObjectId id in srcBtr)
                                {
                                    var sourceViewport = srcTrans.GetObject(id, OpenMode.ForRead, false) as Viewport;
                                    if (sourceViewport != null)
                                    {
                                        if (IsUtilityViewport(sourceViewport, 0.0))
                                        {
                                            psViewportSkippedCount++;
                                            AcadLogger.LogInfo(
                                                $"GD2: Skip default utility viewport before clone: {desiredName} " +
                                                $"handle={sourceViewport.Handle} paperCenter={FormatPoint(sourceViewport.CenterPoint)} " +
                                                $"paperSize=({sourceViewport.Width:F2},{sourceViewport.Height:F2}) " +
                                                $"viewHeight={sourceViewport.ViewHeight:F2} scale={sourceViewport.CustomScale:F8}");
                                            continue;
                                        }

                                        psViewportClonedCount++;
                                    }

                                    psIds.Add(id);
                                }

                                var psIdMap = new IdMapping();
                                int psClonedCount = 0;
                                if (psIds.Count > 0)
                                {
                                    // Dùng Replace thay vì Ignore
                                    sourceDb.WblockCloneObjects(psIds, destBtrId, psIdMap, 
                                        DuplicateRecordCloning.Replace, false);

                                    foreach (ObjectId id in psIds)
                                    {
                                        if (psIdMap.Contains(id) && !psIdMap[id].Value.IsNull)
                                            psClonedCount++;
                                    }

AcadLogger.LogInfo($"GD2: Cloned {psClonedCount}/{psIds.Count} PaperSpace entities to layout '{desiredName}'");
            }
            AcadLogger.LogInfo($"GD2: Included {psViewportClonedCount} source viewport entities, skipped {psViewportSkippedCount} default utility viewport(s); usable source viewport count={sourceViewports.Count}");

            // === GIAI ĐOẠN 2b: Fix cloned viewports after plot settings applied ===
            LogViewportCollection(outputTrans, destBtr, $"DEST after PS clone: {desiredName}");

            // === GIAI ĐOẠN 2c: Fix viewport ViewCenter ===
                                int vpFixedCount = TranslateLayoutViewports(outputTrans, destBtr, desiredName, msOffset);
                                AcadLogger.LogInfo($"GD2: Fixed {vpFixedCount} cloned viewport(s) for '{desiredName}'");
                                LogViewportCollection(outputTrans, destBtr, $"DEST after plot settings: {desiredName}");

                                if (srcStats.EntityCount == 0 && !LayoutHasContent(destBtr, outputTrans))
                                {
                                    AddSchedulePlaceholderContent(destBtr, outputTrans, desiredName, destLayout);
                                }

                                // Lưu thông tin vào danh sách
                                sourceInfos.Add(new SourceFileInfo
                                {
                                    FilePath = source.Path,
                                    LayoutName = desiredName,
                                    MsOffset = msOffset,
                                    MsExtents = srcExtents,
                                    ModelType = modelType,
                                    PlotSettings = savedPlotSettings
                                });

                                // Commit transactions
                                outputTrans.Commit();
                                srcTrans.Commit();

                                clonedCount++;
                                AcadLogger.LogInfo($"Successfully cloned layout '{desiredName}'");
                            }
                        }

                        fileIndex++;
                    }

                    clonedCount += EnsurePendingScheduleOnlyLayouts(outputDb, pendingScheduleOnlyLayouts, sourceInfos);
                    AcadLogger.LogInfo($"Total layouts cloned: {clonedCount}");

                    // Cleanup chỉ xóa các layout mặc định trống
                    CleanupDefaultLayouts(outputDb);
                    ApplyPaperBackgroundPresentation(outputDb, "MultiLayout");
                    RegenerateLayouts(outputDb, "MultiLayout");

                    // Save
                    var dwgVersion = GetDwgVersion(config.DwgVersion);
                    outputDb.SaveAs(config.OutputPath, dwgVersion);
                    AcadLogger.LogInfo($"Saved to: {config.OutputPath}");
                }

                AcadLogger.LogSection("MergeToMultiLayout Complete");
                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"MultiLayout error: {ex.Message}");
                AcadLogger.LogError($"Stack: {ex.StackTrace}");
                return false;
            }
        }

        public bool MergeToSingleLayout(MergeConfig config)
        {
            try
            {
                AcadLogger.Log("[LayoutMerger] Starting SingleLayout merge");

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

var allSourceExtents = new List<Extents3d>();
var allSourceIds = new List<ObjectIdCollection>();
var allSourceDbs = new List<Database>();
var allClonedIds = new List<List<ObjectId>>();

                    using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                    {
                        var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(outputDb);
                        var modelSpace = (BlockTableRecord)outputTrans.GetObject(modelSpaceId, OpenMode.ForWrite);

                        foreach (var source in config.SourceFiles)
                        {
                            if (!File.Exists(source.Path)) continue;

                            AcadLogger.Log($"[LayoutMerger] Processing: {Path.GetFileName(source.Path)}");

                            var sourceDb = new Database(false, true);
                            sourceDb.ReadDwgFile(source.Path, FileShare.ReadWrite, true, "");
                            BindXrefsSafe(sourceDb);

                            using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                            {
                                var sourcePsr = GetSourcePaperSpace(sourceDb, sourceTrans);

                                var ids = new ObjectIdCollection();
                                foreach (ObjectId entId in sourcePsr) ids.Add(entId);

if (ids.Count > 0)
{
var idMap = new IdMapping();
sourceDb.WblockCloneObjects(ids, modelSpaceId, idMap, DuplicateRecordCloning.Ignore, false);

var clonedIds = new List<ObjectId>();
foreach (ObjectId id in ids)
{
if (idMap.Contains(id) && !idMap[id].Value.IsNull)
clonedIds.Add(idMap[id].Value);
}
allClonedIds.Add(clonedIds);

allSourceExtents.Add(GetExtents(sourcePsr));
allSourceIds.Add(ids);
allSourceDbs.Add(sourceDb);
}
                                else
                                {
                                    sourceDb.Dispose();
                                }
                                sourceTrans.Commit();
                            }
                        }

                        double maxGlobalHeight = 0;
                        foreach (var ext in allSourceExtents)
                        {
                            double h = ext.MaxPoint.Y - ext.MinPoint.Y;
                            if (h > maxGlobalHeight) maxGlobalHeight = h;
                        }

                        double xOffset = 0;
                        for (int idx = 0; idx < allSourceIds.Count; idx++)
                        {
                            var ids = allSourceIds[idx];
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

var clonedIds = allClonedIds[idx];
foreach (ObjectId destId in clonedIds)
{
var ent = outputTrans.GetObject(destId, OpenMode.ForWrite) as Entity;
if (ent != null)
{
ent.TransformBy(Matrix3d.Displacement(new Vector3d(xOffset - ext.MinPoint.X, yOffset - ext.MinPoint.Y, 0)));
}
}
                        
                        xOffset += width + LayoutSpacing;
                        AcadLogger.Log($"[LayoutMerger] Offset by ({xOffset}, {yOffset})");
                        }

                        outputTrans.Commit();
                    }

                    for (int i = 0; i < allSourceDbs.Count; i++)
                    {
                        try { allSourceDbs[i].Dispose(); } catch { }
                    }

                    RegenerateLayouts(outputDb, "SingleLayout");

                    var dwgVersion = GetDwgVersion(config.DwgVersion);
                    outputDb.SaveAs(config.OutputPath, dwgVersion);
                    AcadLogger.Log($"[LayoutMerger] SingleLayout completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] SingleLayout error: {ex.Message}");
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
                    int sheetIndex = 0;

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
                                    var placement = new Vector3d(nextSheetX - sourcePaperBounds.MinPoint.X, -sourcePaperBounds.MinPoint.Y, 0.0);

                                    AcadLogger.LogInfo(
                                        $"MODELSPACE placement: label='{label}', bounds={FormatExtents(sourcePaperBounds)}, " +
                                        $"paperExtentEntities={paperExtentEntityCount}, placement={FormatVector(placement)}");

                                    var placedIds = new List<ObjectId>();
                                    int backgroundCount = CreateModelSpaceSheetBackground(
                                        outputDb,
                                        outputTrans,
                                        modelSpace,
                                        sourcePaperBounds,
                                        placedIds);

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
                                        $"bakedModelClones={bakedCount}, moved={movedCount}, finalOriginX={nextSheetX:F2}, " +
                                        $"size=({sheetWidth:F2},{sheetHeight:F2})");

                                    double gap = Math.Max(ModelSpaceSheetMinGap, sheetWidth * 0.05);
                                    nextSheetX += sheetWidth + gap;
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
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Combined DWG index failed: {ex.Message}");
            }
        }

        public void HandleRasterImages(MergeConfig config)
        {
            try
            {
                if (config == null || string.IsNullOrWhiteSpace(config.OutputPath) || !File.Exists(config.OutputPath))
                    return;

                if (!string.Equals(config.RasterImageMode, "EmbedAsOle", StringComparison.OrdinalIgnoreCase))
                {
                    AcadLogger.LogInfo($"Raster image handling skipped: mode={config.RasterImageMode}");
                    return;
                }

                var rasterInfos = ScanRasterImages(config.OutputPath);
                AcadLogger.LogInfo($"Raster image scan: mode={config.RasterImageMode}, count={rasterInfos.Count}");

                if (rasterInfos.Count == 0)
                {
                    AcadLogger.LogInfo("Raster OLE embed requested, but no raster image entities were found.");
                    return;
                }

                foreach (var info in rasterInfos)
                {
                    AcadLogger.LogWarning(
                        $"Raster OLE embed fallback: handle={info.Handle}, owner={info.Owner}, layer={info.Layer}. " +
                        "AutoCAD .NET API did not expose a stable OLE writer in this runtime; keeping raster reference.");
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"Raster image handling failed: {ex.Message}");
            }
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
            var bounds = GetPaperEntityExtentsExcludingViewports(paperSpace, tr, out extentEntities);
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

        private Extents3d GetPaperEntityExtentsExcludingViewports(BlockTableRecord paperSpace, Transaction tr, out int extentsEntityCount)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            extentsEntityCount = 0;

            foreach (ObjectId id in paperSpace)
            {
                try
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased || ent is Viewport)
                        continue;

                    if (string.Equals(ent.Layer, PaperBackgroundLayerName, StringComparison.OrdinalIgnoreCase))
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
                    var orderedLayouts = new List<Tuple<int, string>>();

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
                            int extentsEntityCount = 0;

                            foreach (ObjectId id in paperSpace)
                            {
                                try
                                {
                                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                                    if (ent == null || ent.IsErased)
                                        continue;

                                    paperEntityCount++;
                                    if (ent is Viewport)
                                        viewportCount++;
                                }
                                catch
                                {
                                }
                            }

                            var paperExtents = GetExtents(paperSpace, tr, out extentsEntityCount);
                            layoutInfos.Add(new LayoutRegenInfo
                            {
                                Name = entry.Key,
                                TabOrder = layout.TabOrder,
                                PaperEntityCount = paperEntityCount,
                                ViewportCount = viewportCount,
                                ExtentsEntityCount = extentsEntityCount,
                                PaperExtents = paperExtents,
                                RequiresRegen = viewportCount > 0
                            });
                        }
                        catch
                        {
                            layoutInfos.Add(new LayoutRegenInfo
                            {
                                Name = entry.Key,
                                TabOrder = int.MaxValue,
                                RequiresRegen = true
                            });
                        }
                    }

                    tr.Commit();
                }

                layoutInfos = layoutInfos
                    .OrderBy(x => x.TabOrder)
                    .ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var layoutsToRegen = layoutInfos.Where(x => x.RequiresRegen).ToList();

                AcadLogger.LogSection($"Regenerating Layouts ({mode})");
                AcadLogger.LogInfo("Avoid interacting with AutoCAD until layout regeneration is finished.");
                AcadLogger.LogInfo(
                    $"REGENERATING LAYOUT -> 0/{layoutsToRegen.Count} " +
                    $"(skipped {layoutInfos.Count - layoutsToRegen.Count} stable layout(s) without viewport refresh needs)");

                foreach (var info in layoutInfos.Where(x => !x.RequiresRegen))
                {
                    AcadLogger.LogInfo(
                        $"REGENERATING LAYOUT skip: {info.Name}, " +
                        $"entities={info.PaperEntityCount}, viewports={info.ViewportCount}, " +
                        $"extentsEntities={info.ExtentsEntityCount}, extents={FormatExtents(info.PaperExtents)}");
                }

                if (layoutInfos.Count == 0)
                {
                    AcadLogger.LogWarning($"REGENERATING LAYOUT: no paper layouts found for {mode}");
                    return 0;
                }

                HostApplicationServices.WorkingDatabase = db;

                try
                {
                    db.TileMode = false;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"REGENERATING LAYOUT: failed to switch database to paper space: {ex.Message}");
                }

                if (layoutsToRegen.Count == 0)
                {
                    try
                    {
                        db.UpdateExt(true);
                    }
                    catch (System.Exception ex)
                    {
                        AcadLogger.LogWarning($"REGENERATING LAYOUT: final UpdateExt failed: {ex.Message}");
                    }

                    AcadLogger.LogInfo("REGENERATING LAYOUT complete: 0 layouts required viewport refresh");
                    return 0;
                }

                for (int i = 0; i < layoutsToRegen.Count; i++)
                {
                    var layoutInfo = layoutsToRegen[i];
                    string layoutName = layoutInfo.Name;
                    int paperEntityCount = layoutInfo.PaperEntityCount;
                    int viewportCount = layoutInfo.ViewportCount;
                    int extentsEntityCount = layoutInfo.ExtentsEntityCount;
                    Extents3d paperExtents = layoutInfo.PaperExtents;

                    try
                    {
                        try
                        {
                            LayoutManager.Current.CurrentLayout = layoutName;
                        }
                        catch (System.Exception currentEx)
                        {
                            AcadLogger.LogWarning($"REGENERATING LAYOUT: could not activate '{layoutName}': {currentEx.Message}");
                        }

                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                            if (!layouts.Contains(layoutName))
                            {
                                AcadLogger.LogWarning($"REGENERATING LAYOUT: layout '{layoutName}' disappeared before regen");
                                tr.Commit();
                                continue;
                            }

                            var layout = (Layout)tr.GetObject(layouts.GetAt(layoutName), OpenMode.ForRead);
                            var paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                            paperEntityCount = 0;
                            viewportCount = 0;
                            extentsEntityCount = 0;
                            paperExtents = new Extents3d(Point3d.Origin, Point3d.Origin);

                            foreach (ObjectId id in paperSpace)
                            {
                                try
                                {
                                    var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                                    if (ent == null || ent.IsErased)
                                        continue;

                                    paperEntityCount++;
                                    if (ent is Viewport)
                                        viewportCount++;
                                }
                                catch
                                {
                                }
                            }

                            paperExtents = GetExtents(paperSpace, tr, out extentsEntityCount);
                            tr.Commit();
                        }

                        try
                        {
                            db.UpdateExt(true);
                        }
                        catch (System.Exception updateEx)
                        {
                            AcadLogger.LogWarning($"REGENERATING LAYOUT: UpdateExt failed for '{layoutName}': {updateEx.Message}");
                        }

                        regeneratedCount++;
                        AcadLogger.LogInfo(
                            $"REGENERATING LAYOUT -> {i + 1}/{layoutsToRegen.Count}: {layoutName}, " +
                            $"entities={paperEntityCount}, viewports={viewportCount}, extentsEntities={extentsEntityCount}, " +
                            $"extents={FormatExtents(paperExtents)}");
                    }
                    catch (System.Exception ex)
                    {
                        AcadLogger.LogWarning($"REGENERATING LAYOUT failed for '{layoutName}': {ex.Message}");
                    }
                }

                try
                {
                    db.UpdateExt(true);
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"REGENERATING LAYOUT: final UpdateExt failed: {ex.Message}");
                }

                AcadLogger.LogInfo($"REGENERATING LAYOUT complete: {regeneratedCount}/{layoutsToRegen.Count}");
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
                catch
                {
                }
            }

            return regeneratedCount;
        }

private void BindXrefsSafe(Database db)
{
    try
    {
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

        private bool UseDirectLayoutClone()
        {
            return true;
        }

        private ObjectId CloneLayoutFromSource(Database sourceDb, Database outputDb, Transaction srcTrans, Transaction outputTrans, Layout srcLayout, string desiredName)
        {
            try
            {
                var outputLayouts = (DBDictionary)outputTrans.GetObject(outputDb.LayoutDictionaryId, OpenMode.ForWrite);
                if (outputLayouts.Contains(desiredName))
                {
                    AcadLogger.LogWarning($"CloneLayoutFromSource: layout '{desiredName}' already exists");
                    return ObjectId.Null;
                }

                var layoutForClone = srcLayout;
                var originalName = srcLayout.LayoutName;
                if (!string.Equals(originalName, desiredName, StringComparison.OrdinalIgnoreCase))
                {
                    layoutForClone = (Layout)srcTrans.GetObject(srcLayout.ObjectId, OpenMode.ForWrite);
                    layoutForClone.LayoutName = desiredName;
                    AcadLogger.LogInfo($"GD2: Renamed source layout for clone '{originalName}' -> '{desiredName}'");
                }

                var layoutIds = new ObjectIdCollection { layoutForClone.ObjectId };
                var layoutMap = new IdMapping();
                sourceDb.WblockCloneObjects(layoutIds, outputDb.LayoutDictionaryId, layoutMap, DuplicateRecordCloning.Ignore, false);

                ObjectId destLayoutId = ObjectId.Null;
                if (layoutMap.Contains(layoutForClone.ObjectId) && !layoutMap[layoutForClone.ObjectId].Value.IsNull)
                    destLayoutId = layoutMap[layoutForClone.ObjectId].Value;

                if (destLayoutId.IsNull && outputLayouts.Contains(desiredName))
                    destLayoutId = outputLayouts.GetAt(desiredName);

                if (destLayoutId.IsNull)
                {
                    AcadLogger.LogError($"CloneLayoutFromSource: cloned layout '{desiredName}' was not found in output layout dictionary");
                    return ObjectId.Null;
                }

                var destLayout = (Layout)outputTrans.GetObject(destLayoutId, OpenMode.ForWrite);
                destLayout.LayoutName = desiredName;
                destLayout.TabOrder = outputLayouts.Count - 1;

                var destBtr = (BlockTableRecord)outputTrans.GetObject(destLayout.BlockTableRecordId, OpenMode.ForWrite);
                AcadLogger.LogInfo($"GD2: Layout object cloned '{desiredName}' (BTR={destBtr.ObjectId}, Layout={destLayoutId})");
                return destBtr.ObjectId;
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"CloneLayoutFromSource failed for '{desiredName}': {ex.Message}");
                return ObjectId.Null;
            }
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

                AcadLogger.LogInfo(
                    $"BAKE: Start '{layoutName}' sourceVpCount={sourceViewports.Count}, " +
                    $"sameDb={object.ReferenceEquals(sourceDb, outputDb)}, destPaperSpace={destPaperSpace.ObjectId}");

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
                        sourceTrans, sourceModelSpace, searchExtents, out scannedCount, out noExtentsCount);

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
            Transaction trans,
            BlockTableRecord modelSpace,
            Extents3d searchExtents,
            out int scannedCount,
            out int noExtentsCount)
        {
            var ids = new ObjectIdCollection();
            scannedCount = 0;
            noExtentsCount = 0;

            foreach (ObjectId id in modelSpace)
            {
                scannedCount++;

                try
                {
                    var ent = trans.GetObject(id, OpenMode.ForRead, false) as Entity;
                    if (ent == null || ent.IsErased)
                        continue;

                    var entExtents = ent.GeometricExtents;
                    if (ExtentsIntersect(entExtents, searchExtents))
                        ids.Add(id);
                }
                catch
                {
                    noExtentsCount++;
                }
            }

            return ids;
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
                // Lấy LayoutDictionary và BlockTable
                var layoutDict = (DBDictionary)outputTrans.GetObject(outputDb.LayoutDictionaryId, OpenMode.ForWrite);
                
                // Kiểm tra layout đã tồn tại chưa
                if (layoutDict.Contains(layoutName))
                {
                    AcadLogger.LogWarning($"Layout '{layoutName}' already exists");
                    return ObjectId.Null;
                }
                
                // Tạo BTR mới với tên unique (không copy từ template)
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
                
                // Tạo BTR mới - trống, không copy từ template
                
                // Tạo Layout object mới
                
                // Thêm vào LayoutDictionary
                
                // Link BTR với Layout
                
                // Set TabOrder
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
                layout.CopyFrom(info.PlotSettings);
                layout.BlockTableRecordId = savedBtrId;
                layout.LayoutName = info.LayoutName;
                layout.TabOrder = savedTabOrder;
            }

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

        private void FixViewports(Database outputDb, Transaction outputTrans, Database sourceDb, Transaction sourceTrans,
            ObjectIdCollection srcPsIds, ObjectId destBtrId, IdMapping psIdMap, string layoutName, Vector3d msOffset)
        {
            try
            {
                AcadLogger.LogInfo($"[FixViewports] Starting for '{layoutName}' with MS offset=({msOffset.X:F2}, {msOffset.Y:F2}, {msOffset.Z:F2})");

                // Thu thập viewport từ source
                var srcVps = new List<ViewportInfo>();
                foreach (ObjectId id in srcPsIds)
                {
                    try
                    {
                        var ent = sourceTrans.GetObject(id, OpenMode.ForRead) as Viewport;
                        if (IsModelViewportCandidate(ent))
                        {
                            srcVps.Add(new ViewportInfo
                            {
                                SourceId = id,
                                ViewCenter = ent.ViewCenter,
                                ViewTarget = ent.ViewTarget,
                                ViewDirection = ent.ViewDirection,
                                ViewHeight = ent.ViewHeight,
                                CustomScale = ent.CustomScale,
                                TwistAngle = ent.TwistAngle,
                                CenterPoint = ent.CenterPoint,
                                Width = ent.Width,
                                Height = ent.Height,
                                Number = ent.Number
                            });
                        }
                    }
                    catch { }
                }

                if (srcVps.Count == 0)
                {
                    AcadLogger.LogInfo($"[FixViewports] No viewports found in '{layoutName}'");
                    return;
                }

                int fixedCount = 0;
                foreach (var srcVp in srcVps)
                {
                    try
                    {
                        // Tìm destination viewport qua IdMapping
                        if (!psIdMap.Contains(srcVp.SourceId))
                        {
                            AcadLogger.LogWarning($"[FixViewports] Source viewport {srcVp.SourceId} not in mapping");
                            continue;
                        }

                        ObjectId destId = psIdMap[srcVp.SourceId].Value;
                        if (destId.IsNull)
                        {
                            AcadLogger.LogWarning($"[FixViewports] Destination viewport ID is null");
                            continue;
                        }

                        var destVp = outputTrans.GetObject(destId, OpenMode.ForWrite) as Viewport;
                        if (destVp == null)
                        {
                            AcadLogger.LogWarning($"[FixViewports] Destination object is not a Viewport");
                            continue;
                        }

                        destVp.ViewDirection = srcVp.ViewDirection;
                        destVp.ViewHeight = srcVp.ViewHeight;
                        destVp.CustomScale = srcVp.CustomScale;
                        destVp.TwistAngle = srcVp.TwistAngle;
                        destVp.CenterPoint = srcVp.CenterPoint;
                        destVp.Width = srcVp.Width;
                        destVp.Height = srcVp.Height;
                        destVp.ViewTarget = srcVp.ViewTarget;
                        destVp.On = true;

                        // Keep the Revit DWG target unchanged and pan the model view.
                        var centerOffset = GetViewCenterOffset(destVp, msOffset);
                        destVp.ViewCenter = new Point2d(
                            srcVp.ViewCenter.X + centerOffset.X,
                            srcVp.ViewCenter.Y + centerOffset.Y);

                        fixedCount++;
                        AcadLogger.LogInfo($"[FixViewports] Fixed viewport {destVp.Number}: Scale={srcVp.CustomScale:F4}, Center offset by ({centerOffset.X:F2}, {centerOffset.Y:F2})");
                    }
                    catch (System.Exception ex)
                    {
                        AcadLogger.LogError($"[FixViewports] Error: {ex.Message}");
                    }
                }

                AcadLogger.LogInfo($"[FixViewports] Fixed {fixedCount}/{srcVps.Count} viewports for '{layoutName}'");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"[FixViewports] Error: {ex.Message}");
            }
        }

        private void FixViewportsSimple(Transaction outputTrans, ObjectIdCollection psIds, Vector3d msOffset, string layoutName)
        {
            try
            {
                int fixedCount = 0;
                foreach (ObjectId id in psIds)
                {
                    try
                    {
                        var vp = outputTrans.GetObject(id, OpenMode.ForWrite) as Viewport;
                        if (IsModelViewportCandidate(vp))
                        {
                            var centerOffset = GetViewCenterOffset(vp, msOffset);
                            vp.On = true;
                            vp.ViewCenter = new Point2d(
                                vp.ViewCenter.X + centerOffset.X,
                                vp.ViewCenter.Y + centerOffset.Y);
                            fixedCount++;
                            AcadLogger.LogDebug($"Fixed viewport {vp.Number}: Center adjusted by ({centerOffset.X:F2}, {centerOffset.Y:F2})");
                        }
                    }
                    catch { }
                }
                AcadLogger.LogInfo($"Fixed {fixedCount} viewports for '{layoutName}'");
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"FixViewportsSimple: {ex.Message}");
            }
        }

        private bool LayoutExistsInDb(Database db, string layoutName)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    bool exists = layouts.Contains(layoutName);
                    tr.Commit();
                    return exists;
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"LayoutExistsInDb error: {ex.Message}");
                return false;
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

        private void CopyPlotSettings(Database outputDb, Database sourceDb, Transaction sourceTrans, Layout sourceLayout, string destLayoutName)
        {
            try
            {
                using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                {
                    var layouts = (DBDictionary)outputTrans.GetObject(outputDb.LayoutDictionaryId, OpenMode.ForRead);
                    if (!layouts.Contains(destLayoutName))
                    {
                        AcadLogger.LogWarning($"CopyPlotSettings: Layout '{destLayoutName}' not found");
                        outputTrans.Commit();
                        return;
                    }

                    var destLayoutId = layouts.GetAt(destLayoutName);
                    var destLayout = (Layout)outputTrans.GetObject(destLayoutId, OpenMode.ForWrite);

                    var ps = new PlotSettings(sourceLayout.ModelType);
                    ps.CopyFrom(sourceLayout);
                    destLayout.CopyFrom(ps);

                    AcadLogger.LogInfo($"Copied plot settings for '{destLayoutName}'");
                    outputTrans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"CopyPlotSettings error: {ex.Message}");
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

                        // Chỉ xóa layout có tên mặc định (Layout1, Layout2...) và BTR rỗng
                        // Không xóa layout do user đặt tên
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
                        On = vp.On
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

        private int RecreateLayoutViewports(Transaction trans, BlockTableRecord paperSpace, IReadOnlyList<ViewportInfo> sourceViewports, string layoutName, Vector3d modelOffset)
        {
            int createdCount = 0;

            foreach (var sourceViewport in sourceViewports)
            {
                try
                {
                    var vp = new Viewport();
                    paperSpace.AppendEntity(vp);
                    trans.AddNewlyCreatedDBObject(vp, true);

                    vp.CenterPoint = sourceViewport.CenterPoint;
                    vp.Width = sourceViewport.Width;
                    vp.Height = sourceViewport.Height;
                    vp.ViewDirection = sourceViewport.ViewDirection.Length == 0.0
                        ? Vector3d.ZAxis
                        : sourceViewport.ViewDirection;
                    vp.ViewTarget = sourceViewport.ViewTarget;
                    vp.TwistAngle = sourceViewport.TwistAngle;
                    if (sourceViewport.CustomScale > 0.0)
                        vp.CustomScale = sourceViewport.CustomScale;
                    vp.ViewHeight = sourceViewport.ViewHeight;

                    var centerOffset = GetViewCenterOffset(vp, modelOffset);
                    var newCenter = new Point2d(
                        sourceViewport.ViewCenter.X + centerOffset.X,
                        sourceViewport.ViewCenter.Y + centerOffset.Y);
                    vp.ViewCenter = newCenter;
                    vp.On = true;
                    vp.Locked = sourceViewport.Locked;

                    AcadLogger.LogInfo(
                        $"RECREATE viewport: {layoutName} sourceHandle={sourceViewport.SourceId.Handle} " +
                        $"paperCenter={FormatPoint(sourceViewport.CenterPoint)} paperSize=({sourceViewport.Width:F2},{sourceViewport.Height:F2}) " +
                        $"sourceCenter={FormatPoint(sourceViewport.ViewCenter)} newCenter={FormatPoint(newCenter)} " +
                        $"centerDelta=({centerOffset.X:F4},{centerOffset.Y:F4}) modelOffset={FormatVector(modelOffset)} " +
                        $"sourceOn={sourceViewport.On} newOn={vp.On} locked={vp.Locked}");
                    LogViewportState($"RECREATE after: {layoutName}", vp);
                    createdCount++;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"RECREATE viewport error: {layoutName}: {ex.Message}");
                }
            }

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

        private int TranslateLayoutViewports(Transaction trans, BlockTableRecord paperSpace, string layoutName, Vector3d modelOffset)
        {
            int fixedCount = 0;
            int skippedCount = 0;
            int erasedUtilityCount = 0;
            var viewportIds = new List<ObjectId>();
            double maxViewArea = 0.0;

            foreach (ObjectId id in paperSpace)
                viewportIds.Add(id);

            foreach (ObjectId id in viewportIds)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForRead, false) as Viewport;
                    if (IsRawModelViewport(vp))
                        maxViewArea = Math.Max(maxViewArea, GetViewportViewArea(vp));
                }
                catch { }
            }

            foreach (ObjectId id in viewportIds)
            {
                try
                {
                    var vp = trans.GetObject(id, OpenMode.ForWrite, false) as Viewport;
                    if (vp == null)
                        continue;

                    if (IsUtilityViewport(vp, maxViewArea))
                    {
                        AcadLogger.LogInfo($"FIX erase utility viewport: {layoutName} VP#{vp.Number} handle={vp.Handle} paperSize=({vp.Width:F2},{vp.Height:F2}) viewHeight={vp.ViewHeight:F2} scale={vp.CustomScale:F8} viewArea={GetViewportViewArea(vp):F2} maxArea={maxViewArea:F2}");
                        vp.Erase();
                        erasedUtilityCount++;
                        continue;
                    }

                    if (!IsModelViewportCandidate(vp))
                    {
                        skippedCount++;
                        AcadLogger.LogDebug($"FIX skip: {layoutName} VP#{vp.Number} handle={vp.Handle} is not a model viewport candidate");
                        continue;
                    }

                    LogViewportState($"FIX before: {layoutName}", vp);

                    var oldCenter = vp.ViewCenter;
                    var oldTarget = vp.ViewTarget;
                    bool oldOn = vp.On;
                    var centerOffset = GetViewCenterOffset(vp, modelOffset);
                    bool wasLocked = vp.Locked;

                    if (wasLocked)
                        vp.Locked = false;

                    var newCenter = new Point2d(
                        vp.ViewCenter.X + centerOffset.X,
                        vp.ViewCenter.Y + centerOffset.Y);
                    vp.ViewCenter = newCenter;

                    try
                    {
                        vp.On = true;
                    }
                    catch (System.Exception onEx)
                    {
                        AcadLogger.LogWarning($"FIX viewport on-state error: {layoutName} VP#{vp.Number} handle={vp.Handle}: {onEx.Message}");
                    }

                    if (wasLocked)
                        vp.Locked = true;

                    AcadLogger.LogInfo(
                        $"FIX delta: {layoutName} VP#{vp.Number} " +
                        $"oldCenter={FormatPoint(oldCenter)} newCenter={FormatPoint(newCenter)} " +
                        $"centerDelta=({centerOffset.X:F4},{centerOffset.Y:F4}) " +
                        $"targetKept={FormatPoint(oldTarget)} modelOffset={FormatVector(modelOffset)} on={oldOn}->{vp.On} " +
                        $"visibleBefore={FormatExtents(GetViewportViewExtents(oldCenter, vp))} " +
                        $"visibleAfter={FormatExtents(GetViewportViewExtents(newCenter, vp))}");

                    LogViewportState($"FIX after: {layoutName}", vp);
                    fixedCount++;
                }
                catch (System.Exception ex)
                {
                    AcadLogger.LogWarning($"FIX viewport error: {layoutName} {id}: {ex.Message}");
                }
            }

            AcadLogger.LogInfo($"FIX summary: {layoutName} fixed={fixedCount}, skipped={skippedCount}, erasedUtility={erasedUtilityCount}, modelOffset={FormatVector(modelOffset)}");
            return fixedCount;
        }

        private bool IsModelViewportCandidate(Viewport vp)
        {
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

            return defaultPaperViewport;
        }

        private double GetViewportViewArea(Viewport vp)
        {
            if (vp == null || vp.Width <= 0.0 || vp.Height <= 0.0 || vp.ViewHeight <= 0.0)
                return 0.0;

            return vp.ViewHeight * (vp.ViewHeight * (vp.Width / vp.Height));
        }

        private Vector2d GetViewCenterOffset(Viewport vp, Vector3d modelOffset)
        {
            try
            {
                var wcsToDcs = GetWorldToDcsTransform(vp);
                var origin = Point3d.Origin.TransformBy(wcsToDcs);
                var offsetPoint = new Point3d(modelOffset.X, modelOffset.Y, modelOffset.Z).TransformBy(wcsToDcs);
                var delta = origin.GetVectorTo(offsetPoint);
                return new Vector2d(delta.X, delta.Y);
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"GetViewCenterOffset fallback: {ex.Message}");
                return new Vector2d(modelOffset.X, modelOffset.Y);
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

        private Extents3d GetExtents(BlockTableRecord btr)
        {
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

            if (minX == double.MaxValue)
                return new Extents3d(new Point3d(0, 0, 0), new Point3d(0, 0, 0));

            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        }

        private Extents3d GetExtents(BlockTableRecord btr, Transaction trans, out int extentsEntityCount)
        {
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

            if (minX == double.MaxValue)
                return new Extents3d(new Point3d(0, 0, 0), new Point3d(0, 0, 0));

            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
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
        }
    }
}
