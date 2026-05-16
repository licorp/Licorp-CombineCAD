using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;

namespace Licorp_MergeSheets
{
    public class LayerMappingService
    {
        public int ApplyLayerMapping(Database db, Transaction tr, List<LayerMappingRule> rules)
        {
            if (db == null || rules == null || rules.Count == 0)
                return 0;

            int mappedCount = 0;

            try
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                var targetLayers = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);

                foreach (var rule in rules)
                {
                    if (string.IsNullOrWhiteSpace(rule.TargetLayer))
                        continue;

                    if (!layerTable.Has(rule.TargetLayer))
                    {
                        layerTable.UpgradeOpen();
                        var newLayer = new LayerTableRecord
                        {
                            Name = rule.TargetLayer
                        };
                        ObjectId layerId = layerTable.Add(newLayer);
                        tr.AddNewlyCreatedDBObject(newLayer, true);
                        targetLayers[rule.TargetLayer] = layerId;
                        AcadLogger.LogInfo($"LayerMappingService: Created target layer '{rule.TargetLayer}'");
                    }
                    else
                    {
                        targetLayers[rule.TargetLayer] = layerTable[rule.TargetLayer];
                    }
                }

                var modelSpace = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                var allBtrs = new List<BlockTableRecord>();
                allBtrs.Add(modelSpace);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    if (string.Equals(entry.Key, "Model", StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        if (layout.BlockTableRecordId.IsNull)
                            continue;

                        var psBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                        allBtrs.Add(psBtr);
                    }
                    catch { }
                }

                foreach (var btr in allBtrs)
                {
                    foreach (ObjectId entId in btr)
                    {
                        try
                        {
                            var ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                            if (ent == null || ent.IsErased)
                                continue;

                            string currentLayer = ent.Layer;
                            string mappedLayer = FindMatchingLayer(currentLayer, rules);

                            if (mappedLayer != null && !string.Equals(currentLayer, mappedLayer, StringComparison.OrdinalIgnoreCase))
                            {
                                ent.Layer = mappedLayer;
                                mappedCount++;
                            }
                        }
                        catch { }
                    }
                }

                AcadLogger.LogInfo($"LayerMappingService: Remapped {mappedCount} entity(ies) across {allBtrs.Count} block(s)");
            }
            catch (Exception ex)
            {
                AcadLogger.LogError($"LayerMappingService: Failed: {ex.Message}");
            }

            return mappedCount;
        }

        private string FindMatchingLayer(string currentLayer, List<LayerMappingRule> rules)
        {
            foreach (var rule in rules)
            {
                if (string.IsNullOrWhiteSpace(rule.SourcePattern))
                    continue;

                if (rule.IsWildcard)
                {
                    if (LikeOperator(currentLayer, rule.SourcePattern))
                        return rule.TargetLayer;
                }
                else
                {
                    if (string.Equals(currentLayer, rule.SourcePattern, StringComparison.OrdinalIgnoreCase))
                        return rule.TargetLayer;
                }
            }

            return null;
        }

        private bool LikeOperator(string input, string pattern)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(pattern))
                return false;

            string regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";

            return System.Text.RegularExpressions.Regex.IsMatch(input, regexPattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
    }
}
