using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Diagnostics;

namespace Licorp_MergeSheets
{
    public class LayoutMerger
    {
        private const double LayoutSpacing = 50.0;

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
                        AcadLogger.LogInfo($"Using base file: {baseFile}");
                        break;
                    }
                }

                if (baseFile == null)
                {
                    AcadLogger.LogError("No valid source files found");
                    return false;
                }

                var outputDb = new Database(false, true);

                using (outputDb)
                {
                    outputDb.ReadDwgFile(baseFile, FileShare.ReadWrite, true, "");
                    AcadLogger.LogInfo("Base file opened successfully");

                    int clonedCount = 0;
                    int layoutIndex = 1;

                    foreach (var source in config.SourceFiles)
                    {
                        if (!File.Exists(source.Path))
                        {
                            AcadLogger.LogWarning($"Source not found: {source.Path}");
                            continue;
                        }

                        if (source.Path.Equals(baseFile, StringComparison.OrdinalIgnoreCase))
                        {
                            var desiredName = source.Layout;
                            if (!string.IsNullOrEmpty(desiredName))
                            {
                                RenameLayoutInDb(outputDb, "Layout1", desiredName);
                                clonedCount++;
                                AcadLogger.LogInfo($"Renamed base layout to '{desiredName}'");
                            }
                            continue;
                        }

                        AcadLogger.LogInfo($"Processing: {Path.GetFileName(source.Path)}");

                        var sourceDb = new Database(false, true);
                        using (sourceDb)
                        {
sourceDb.ReadDwgFile(source.Path, FileShare.Read, true, "");

                            try
                            {
                                var xrefIds = new ObjectIdCollection();
                                sourceDb.BindXrefs(xrefIds, true);
                            }
                            catch (Exception ex) { AcadLogger.LogWarning($"BindXrefs failed: {ex.Message}"); }

                            using (var srcTrans = sourceDb.TransactionManager.StartTransaction())
                            {
                                var srcLayoutDict = (DBDictionary)srcTrans.GetObject(
                                    sourceDb.LayoutDictionaryId, OpenMode.ForRead);

                                foreach (DBDictionaryEntry entry in srcLayoutDict)
                                {
                                    if (entry.Key == "Model")
                                        continue;

                                    string desiredName = source.Layout ?? entry.Key;
                                    string sourceLayoutName = entry.Key;

                                    if (LayoutExistsInDb(outputDb, desiredName))
                                    {
                                        AcadLogger.LogWarning($"Layout '{desiredName}' already exists, skipping");
                                        continue;
                                    }

                                    var idsToClone = new ObjectIdCollection { entry.Value };
                                    var mapping = new IdMapping();

                                    sourceDb.WblockCloneObjects(
                                        idsToClone,
                                        outputDb.LayoutDictionaryId,
                                        mapping,
                                        DuplicateRecordCloning.Ignore,
                                        false);

                                    if (desiredName != sourceLayoutName)
                                    {
                                        RenameClonedLayout(outputDb, mapping, sourceLayoutName, desiredName);
                                    }

                                    CopyPlotSettings(outputDb, sourceDb, srcTrans, entry.Value, desiredName);

                                    clonedCount++;
                                    AcadLogger.LogInfo($"Cloned layout '{desiredName}' from {Path.GetFileName(source.Path)}");
                                    break;
                                }

                                srcTrans.Commit();
                            }
                        }

                        layoutIndex++;
                    }

                    AcadLogger.LogInfo($"Total layouts cloned: {clonedCount}");

                    CleanupDefaultLayouts(outputDb);

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
                    outputDb.ReadDwgFile(config.OutputPath, FileShare.ReadWrite, true, "");

                    var allSourceExtents = new List<Extents3d>();
                    var allSourceIds = new List<ObjectIdCollection>();
                    var allSourceDbs = new List<Database>();

                    using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                    {
                        var outputBt = (BlockTable)outputTrans.GetObject(outputDb.BlockTableId, OpenMode.ForRead);
                        var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(outputDb);
                        var modelSpace = (BlockTableRecord)outputTrans.GetObject(modelSpaceId, OpenMode.ForWrite);

                        foreach (var source in config.SourceFiles)
                        {
                            if (!File.Exists(source.Path))
                                continue;

                            AcadLogger.Log($"[LayoutMerger] Processing: {Path.GetFileName(source.Path)}");

                            var sourceDb = new Database(false, true);
                            sourceDb.ReadDwgFile(source.Path, FileShare.Read, true, "");

                            try { var xrefIds = new ObjectIdCollection(); sourceDb.BindXrefs(xrefIds, true); }
                            catch (Exception ex) { AcadLogger.LogWarning($"BindXrefs failed: {ex.Message}"); }

                            using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                            {
                                var sourcePsr = GetSourcePaperSpace(sourceDb, sourceTrans);

                                var ids = new ObjectIdCollection();
                                foreach (ObjectId entId in sourcePsr)
                                {
                                    ids.Add(entId);
                                }

                                if (ids.Count > 0)
                                {
                                    var idMap = new IdMapping();
                                    sourceDb.WblockCloneObjects(ids, modelSpaceId, idMap, DuplicateRecordCloning.Replace, false);

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

                            foreach (ObjectId id in ids)
                            {
                                var ent = (Entity)outputTrans.GetObject(id, OpenMode.ForWrite);
                                if (ent != null)
                                {
                                    ent.TransformBy(Matrix3d.Displacement(new Vector3d(
                                        xOffset - ext.MinPoint.X,
                                        yOffset - ext.MinPoint.Y,
                                        0)));
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
                AcadLogger.Log("[LayoutMerger] Starting ModelSpace merge");

                var outputDb = new Database(false, true);

                using (outputDb)
                {
                    outputDb.ReadDwgFile(config.OutputPath, FileShare.ReadWrite, true, "");

                    using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                    {
                        var modelSpaceId = SymbolUtilityServices.GetBlockModelSpaceId(outputDb);
                        var modelSpace = (BlockTableRecord)outputTrans.GetObject(modelSpaceId, OpenMode.ForWrite);

                        int cols = (int)Math.Ceiling(Math.Sqrt(config.SourceFiles.Count));
                        int row = 0, col = 0;

                        foreach (var source in config.SourceFiles)
                        {
                            if (!File.Exists(source.Path))
                                continue;

                            AcadLogger.Log($"[LayoutMerger] Processing: {Path.GetFileName(source.Path)}");

                            var sourceDb = new Database(false, true);
                            using (sourceDb)
                            {
                                sourceDb.ReadDwgFile(source.Path, FileShare.Read, true, "");

                                try { var xrefIds = new ObjectIdCollection(); sourceDb.BindXrefs(xrefIds, true); }
                                catch (Exception ex) { AcadLogger.LogWarning($"BindXrefs failed: {ex.Message}"); }

                                using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                                {
                                    var sourcePsr = GetSourcePaperSpace(sourceDb, sourceTrans);

                                    var ids = new ObjectIdCollection();
                                    foreach (ObjectId entId in sourcePsr)
                                    {
                                        ids.Add(entId);
                                    }

                                    if (ids.Count > 0)
                                    {
                                        var idMap = new IdMapping();
                                        sourceDb.WblockCloneObjects(ids, modelSpaceId, idMap, DuplicateRecordCloning.Replace, false);

                                        var ext = GetExtents(sourcePsr);
                                        double xOffset = col * (ext.MaxPoint.X - ext.MinPoint.X + LayoutSpacing * 2);
                                        double yOffset = -row * (ext.MaxPoint.Y - ext.MinPoint.Y + LayoutSpacing * 2);

                                        foreach (ObjectId id in ids)
                                        {
                                            var ent = (Entity)outputTrans.GetObject(id, OpenMode.ForWrite);
                                            if (ent != null)
                                            {
                                                ent.TransformBy(Matrix3d.Displacement(new Vector3d(xOffset - ext.MinPoint.X, yOffset - ext.MinPoint.Y, 0)));
                                            }
                                        }

                                        AcadLogger.Log($"[LayoutMerger] Placed at ({xOffset}, {yOffset})");
                                    }

                                    sourceTrans.Commit();
                                }
                            }

                            col++;
                            if (col >= cols)
                            {
                                col = 0;
                                row++;
                            }
                        }

                        outputTrans.Commit();
                    }

                    var dwgVersion = GetDwgVersion(config.DwgVersion);
                    outputDb.SaveAs(config.OutputPath, dwgVersion);
                    AcadLogger.Log($"[LayoutMerger] ModelSpace completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] ModelSpace error: {ex.Message}");
                return false;
            }
        }

        private void CopyPlotSettings(Database destDb, Database srcDb, Transaction srcTrans, ObjectId srcLayoutId, string destLayoutName)
        {
            try
            {
                var srcLayout = (Layout)srcTrans.GetObject(srcLayoutId, OpenMode.ForRead);

                using (var destTr = destDb.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)destTr.GetObject(destDb.LayoutDictionaryId, OpenMode.ForRead);

                    if (layoutDict.Contains(destLayoutName))
                    {
                        var destLayout = (Layout)destTr.GetObject(layoutDict.GetAt(destLayoutName), OpenMode.ForWrite);
                        destLayout.CopyFrom(srcLayout);
                        AcadLogger.LogInfo($"Copied plot settings for '{destLayoutName}'");
                    }

                    destTr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogWarning($"CopyPlotSettings error: {ex.Message}");
            }
        }

        private void EnsureLayoutNameAvailable(Database db, string name)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite);
                    if (layoutDict.Contains(name))
                    {
                        layoutDict.Remove(name);
                        AcadLogger.Log($"[LayoutMerger] Removed existing layout '{name}'");
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] EnsureLayoutNameAvailable error: {ex.Message}");
            }
        }

        private void RenameClonedLayout(Database db, IdMapping mapping, string oldName, string newName)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    if (layoutDict.Contains(oldName))
                    {
                        var layoutId = layoutDict.GetAt(oldName);
                        var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForWrite);
                        layout.LayoutName = newName;
                        AcadLogger.Log($"[LayoutMerger] Renamed layout '{oldName}' to '{newName}'");
                    }
                    else if (layoutDict.Contains(newName))
                    {
                        AcadLogger.Log($"[LayoutMerger] Layout already named '{newName}'");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                AcadLogger.Log($"[LayoutMerger] RenameClonedLayout error: {ex.Message}");
            }
        }

        private bool LayoutExistsInDb(Database db, string layoutName)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    return layoutDict.Contains(layoutName);
                }
            }
            catch
            {
                return false;
            }
        }

        private void RenameLayoutInDb(Database db, string oldName, string newName)
        {
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    if (layoutDict.Contains(oldName))
                    {
                        var layoutId = layoutDict.GetAt(oldName);
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
                        AcadLogger.Log($"[LayoutMerger] Deleted extra layout '{name}'");
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

        private DwgVersion GetDwgVersion(string version)
        {
            if (string.IsNullOrEmpty(version) || version == "Current")
                return DwgVersion.Current;

            switch (version)
            {
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
    }
}