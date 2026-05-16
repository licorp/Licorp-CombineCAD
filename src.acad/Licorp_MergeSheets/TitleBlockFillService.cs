using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace Licorp_MergeSheets
{
    public class TitleBlockFillService
    {
        public int FillTitleBlocks(string dwgPath, List<TitleBlockFieldMapping> mappings)
        {
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
            {
                AcadLogger.LogWarning("TitleBlockFillService: DWG file not found");
                return 0;
            }

            if (mappings == null || mappings.Count == 0)
            {
                AcadLogger.LogWarning("TitleBlockFillService: No field mappings provided");
                return 0;
            }

            int filledCount = 0;

            try
            {
                var db = new Database(false, true);
                using (db)
                {
                    db.ReadDwgFile(dwgPath, FileShare.ReadWrite, true, "");
                    db.CloseInput(true);

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                        foreach (var mapping in mappings)
                        {
                            try
                            {
                                if (string.IsNullOrWhiteSpace(mapping.LayoutName))
                                    continue;

                                if (!layoutDict.Contains(mapping.LayoutName))
                                {
                                    AcadLogger.LogWarning($"TitleBlockFillService: Layout '{mapping.LayoutName}' not found");
                                    continue;
                                }

                                var layoutId = layoutDict.GetAt(mapping.LayoutName);
                                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                                foreach (ObjectId entId in btr)
                                {
                                    try
                                    {
                                        var br = tr.GetObject(entId, OpenMode.ForRead, false) as BlockReference;
                                        if (br == null)
                                            continue;

                                        if (!string.IsNullOrEmpty(mapping.BlockName) &&
                                            !string.Equals(br.Name, mapping.BlockName, StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        var attrDict = br.AttributeCollection;
                                        if (attrDict == null)
                                            continue;

                                        foreach (ObjectId attrId in attrDict)
                                        {
                                            var attrRef = tr.GetObject(attrId, OpenMode.ForWrite, false) as AttributeReference;
                                            if (attrRef == null)
                                                continue;

                                            if (string.Equals(attrRef.Tag, mapping.AttributeTag, StringComparison.OrdinalIgnoreCase))
                                            {
                                                attrRef.TextString = mapping.Value;
                                                filledCount++;
                                                AcadLogger.LogInfo($"TitleBlockFillService: Set '{mapping.AttributeTag}' = '{mapping.Value}' in layout '{mapping.LayoutName}'");
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch (Exception ex)
                            {
                                AcadLogger.LogWarning($"TitleBlockFillService: Error processing mapping for layout '{mapping.LayoutName}': {ex.Message}");
                            }
                        }

                        tr.Commit();
                    }

                    var dwgVersion = DwgVersion.Current;
                    db.SaveAs(dwgPath, dwgVersion);
                    AcadLogger.LogInfo($"TitleBlockFillService: Filled {filledCount} attribute(s) and saved");
                }
            }
            catch (Exception ex)
            {
                AcadLogger.LogError($"TitleBlockFillService: Failed: {ex.Message}");
            }

            return filledCount;
        }

        public List<TitleBlockFieldMapping> LoadMappingsFromCsv(string csvPath, string defaultAttributeName = null)
        {
            var mappings = new List<TitleBlockFieldMapping>();

            if (string.IsNullOrWhiteSpace(csvPath) || !File.Exists(csvPath))
            {
                AcadLogger.LogWarning($"TitleBlockFillService: CSV file not found: {csvPath}");
                return mappings;
            }

            try
            {
                var lines = File.ReadAllLines(csvPath);
                if (lines.Length < 2)
                {
                    AcadLogger.LogWarning("TitleBlockFillService: CSV file is empty or has no data rows");
                    return mappings;
                }

                var headers = ParseCsvLine(lines[0]);
                int layoutCol = FindColumnIndex(headers, "Layout");
                int blockCol = FindColumnIndex(headers, "Block");
                int tagCol = FindColumnIndex(headers, "Tag");
                int valueCol = FindColumnIndex(headers, "Value");

                if (tagCol < 0 || valueCol < 0)
                {
                    AcadLogger.LogWarning("TitleBlockFillService: CSV must have 'Tag' and 'Value' columns");
                    return mappings;
                }

                for (int i = 1; i < lines.Length; i++)
                {
                    var line = lines[i].Trim();
                    if (string.IsNullOrEmpty(line))
                        continue;

                    var fields = ParseCsvLine(line);
                    if (fields.Count <= valueCol)
                        continue;

                    var mapping = new TitleBlockFieldMapping
                    {
                        LayoutName = layoutCol >= 0 && layoutCol < fields.Count ? fields[layoutCol].Trim() : null,
                        BlockName = blockCol >= 0 && blockCol < fields.Count ? fields[blockCol].Trim() : null,
                        AttributeTag = fields[tagCol].Trim(),
                        Value = fields[valueCol].Trim()
                    };

                    if (!string.IsNullOrEmpty(mapping.AttributeTag))
                        mappings.Add(mapping);
                }

                AcadLogger.LogInfo($"TitleBlockFillService: Loaded {mappings.Count} mapping(s) from CSV");
            }
            catch (Exception ex)
            {
                AcadLogger.LogError($"TitleBlockFillService: Failed to load CSV: {ex.Message}");
            }

            return mappings;
        }

        private List<string> ParseCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            string current = "";

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(current);
                    current = "";
                }
                else
                {
                    current += c;
                }
            }

            fields.Add(current);
            return fields;
        }

        private int FindColumnIndex(List<string> headers, string name)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (string.Equals(headers[i].Trim(), name, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }
    }
}
