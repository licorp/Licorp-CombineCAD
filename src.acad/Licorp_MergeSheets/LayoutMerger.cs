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
                Debug.WriteLine("[Merger] Starting MultiLayout merge");

                var outputDb = new Database(false, true);

                using (outputDb)
                {
                    outputDb.ReadDwgFile(config.OutputPath, FileShare.ReadWrite, true, "");

                    using (var outputTrans = outputDb.TransactionManager.StartTransaction())
                    {
                        var outputBt = (BlockTable)outputTrans.GetObject(outputDb.BlockTableId, OpenMode.ForRead);

                        int layoutIndex = 1;
                        foreach (var source in config.SourceFiles)
                        {
                            if (!File.Exists(source.Path))
                            {
                                Debug.WriteLine($"[Merger] Source not found: {source.Path}");
                                continue;
                            }

                            Debug.WriteLine($"[Merger] Processing: {Path.GetFileName(source.Path)}");

                            var sourceDb = new Database(false, true);
                            using (sourceDb)
                            {
                                sourceDb.ReadDwgFile(source.Path, FileShare.Read, true, "");

using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                    {
                    var sourcePsr = GetSourcePaperSpace(sourceDb, sourceTrans);

                    string layoutName = source.Layout ?? $"Layout{layoutIndex}";

                                    var destPsrId = CreateLayoutInCurrentDb(outputDb, outputTrans, layoutName);

                                    var ids = new ObjectIdCollection();
                                    foreach (ObjectId entId in sourcePsr)
                                    {
                                        ids.Add(entId);
                                    }

                                    if (ids.Count > 0)
                                    {
                                        var idMap = new IdMapping();
                                        sourceDb.WblockCloneObjects(ids, destPsrId, idMap, DuplicateRecordCloning.Replace, false);
                                        Debug.WriteLine($"[Merger] Cloned {ids.Count} entities to layout {layoutName}");
                                    }

                                    sourceTrans.Commit();
                                }
                            }

                            layoutIndex++;
                        }

                        outputTrans.Commit();
                    }

                    outputDb.SaveAs(config.OutputPath, DwgVersion.Current);
                    Debug.WriteLine($"[Merger] MultiLayout completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"[Merger] MultiLayout error: {ex.Message}");
                return false;
            }
        }

        public bool MergeToSingleLayout(MergeConfig config)
        {
            try
            {
                Debug.WriteLine("[Merger] Starting SingleLayout merge");

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

                            Debug.WriteLine($"[Merger] Processing: {Path.GetFileName(source.Path)}");

                            var sourceDb = new Database(false, true);
                            sourceDb.ReadDwgFile(source.Path, FileShare.Read, true, "");

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
                            Debug.WriteLine($"[Merger] Offset by ({xOffset}, {yOffset})");
                        }

                        outputTrans.Commit();
                    }

                    for (int i = 0; i < allSourceDbs.Count; i++)
                    {
                        try { allSourceDbs[i].Dispose(); } catch { }
                    }

                    outputDb.SaveAs(config.OutputPath, DwgVersion.Current);
                    Debug.WriteLine($"[Merger] SingleLayout completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"[Merger] SingleLayout error: {ex.Message}");
                return false;
            }
        }

        public bool MergeToModelSpace(MergeConfig config)
        {
            try
            {
                Debug.WriteLine("[Merger] Starting ModelSpace merge");

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

                            Debug.WriteLine($"[Merger] Processing: {Path.GetFileName(source.Path)}");

                            var sourceDb = new Database(false, true);
                            using (sourceDb)
                            {
                                sourceDb.ReadDwgFile(source.Path, FileShare.Read, true, "");

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

                        Debug.WriteLine($"[Merger] Placed at ({xOffset}, {yOffset})");
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

                    outputDb.SaveAs(config.OutputPath, DwgVersion.Current);
                    Debug.WriteLine($"[Merger] ModelSpace completed: {config.OutputPath}");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"[Merger] ModelSpace error: {ex.Message}");
                return false;
            }
        }

        private ObjectId CreateLayoutInCurrentDb(Database db, Transaction trans, string layoutName)
        {
            var lt = (LayoutManager)LayoutManager.Current;

            var layoutId = lt.CreateLayout(layoutName);
            lt.CurrentLayout = layoutName;

            var layout = (Layout)trans.GetObject(layoutId, OpenMode.ForRead);
            return layout.BlockTableRecordId;
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
                Debug.WriteLine($"[Merger] Error getting Paper Space: {ex.Message}");
            }

            Debug.WriteLine("[Merger] Fallback to Model Space");
            return (BlockTableRecord)sourceTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(sourceDb), OpenMode.ForRead);
        }

    private Extents3d GetExtents(BlockTableRecord btr)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (ObjectId id in btr)
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

            if (minX == double.MaxValue)
                return new Extents3d(new Point3d(0, 0, 0), new Point3d(0, 0, 0));

            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        }
    }
}