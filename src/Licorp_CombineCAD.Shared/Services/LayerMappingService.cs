using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Licorp_CombineCAD.Models;

namespace Licorp_CombineCAD.Services
{
    public class LayerMappingService
    {
        private readonly Document _document;

        public LayerMappingService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public bool ExportLayerMapping(string setupName, string filePath)
        {
            if (string.IsNullOrEmpty(setupName) || string.IsNullOrEmpty(filePath))
                return false;

            try
            {
                var dwgSettings = new FilteredElementCollector(_document)
                    .OfClass(typeof(ExportDWGSettings))
                    .Cast<ExportDWGSettings>()
                    .FirstOrDefault(s => s.Name == setupName);

                if (dwgSettings == null)
                {
                    Trace.WriteLine($"[LayerMapping] Setup '{setupName}' not found");
                    return false;
                }

                var options = dwgSettings.GetDWGExportOptions();
                var entries = new List<string> { "Category\tSubCategory\tLayerName\tColorId\tLinetype" };

                var layerTable = GetExportLayerTable(options);
                if (layerTable != null)
                {
                    var tableType = layerTable.GetType();

                    var countProp = tableType.GetProperty("Count");
                    int count = countProp != null ? (int)countProp.GetValue(layerTable) : 0;

                    var getItemMethod = tableType.GetProperty("Item", new[] { typeof(int) });
                    var enumeratorMethod = tableType.GetMethod("GetEnumerator");

                    if (getItemMethod != null && count > 0)
                    {
                        for (int i = 0; i < count; i++)
                        {
                            try
                            {
                                var entry = getItemMethod.GetValue(layerTable, new object[] { i });
                                if (entry != null)
                                {
                                    var line = ExtractLayerEntry(entry);
                                    if (line != null)
                                        entries.Add(line);
                                }
                            }
                            catch (Exception ex)
                            {
                                Trace.WriteLine($"[LayerMapping] Error reading entry {i}: {ex.Message}");
                            }
                        }
                    }
                    else if (enumeratorMethod != null)
                    {
                        try
                        {
                            var enumerator = enumeratorMethod.Invoke(layerTable, null);
                            var moveNext = enumerator.GetType().GetMethod("MoveNext");
                            var currentProp = enumerator.GetType().GetProperty("Current");

                            while ((bool)moveNext.Invoke(enumerator, null))
                            {
                                var entry = currentProp.GetValue(enumerator);
                                if (entry != null)
                                {
                                    var line = ExtractLayerEntry(entry);
                                    if (line != null)
                                        entries.Add(line);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Trace.WriteLine($"[LayerMapping] Enumerator error: {ex.Message}");
                        }
                    }

                    Trace.WriteLine($"[LayerMapping] Exported {entries.Count - 1} entries");
                }
                else
                {
                    Trace.WriteLine("[LayerMapping] ExportLayerTable not available, exporting header only");
                }

                File.WriteAllLines(filePath, entries);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LayerMapping] Export error: {ex.Message}");
                return false;
            }
        }

        public bool ImportLayerMapping(string setupName, string filePath)
        {
            if (string.IsNullOrEmpty(setupName) || string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                var dwgSettings = new FilteredElementCollector(_document)
                    .OfClass(typeof(ExportDWGSettings))
                    .Cast<ExportDWGSettings>()
                    .FirstOrDefault(s => s.Name == setupName);

                if (dwgSettings == null)
                {
                    Trace.WriteLine($"[LayerMapping] Setup '{setupName}' not found");
                    return false;
                }

                var options = dwgSettings.GetDWGExportOptions();
                var lines = File.ReadAllLines(filePath);
                if (lines.Length <= 1) return true;

                var importedEntries = new List<LayerMappingEntry>();
                for (int i = 1; i < lines.Length; i++)
                {
                    var entry = LayerMappingEntry.Parse(lines[i]);
                    if (entry != null)
                        importedEntries.Add(entry);
                }

                var layerTable = GetExportLayerTable(options);
                if (layerTable != null)
                {
                    var tableType = layerTable.GetType();
                    var addMethod = tableType.GetMethod("Add");
                    var clearMethod = tableType.GetMethod("Clear");

                    if (clearMethod != null)
                    {
                        try { clearMethod.Invoke(layerTable, null); }
                        catch { }
                    }

                    if (addMethod != null)
                    {
                        var entryType = addMethod.GetParameters().FirstOrDefault()?.ParameterType;
                        if (entryType != null)
                        {
                            foreach (var imported in importedEntries)
                            {
                                try
                                {
                                    var newEntry = CreateLayerEntry(entryType, imported);
                                    if (newEntry != null)
                                        addMethod.Invoke(layerTable, new[] { newEntry });
                                }
                                catch (Exception ex)
                                {
                                    Trace.WriteLine($"[LayerMapping] Error adding entry: {ex.Message}");
                                }
                            }
                        }
                    }

                    var setLayerTableMethod = options.GetType().GetMethod("SetExportLayerTable");
                    if (setLayerTableMethod != null)
                    {
                        setLayerTableMethod.Invoke(options, new[] { layerTable });
                    }

                    Trace.WriteLine($"[LayerMapping] Imported {importedEntries.Count} entries");
                }
                else
                {
                    Trace.WriteLine("[LayerMapping] ExportLayerTable not available for import");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LayerMapping] Import error: {ex.Message}");
                return false;
            }
        }

        public bool ValidateFile(string filePath, out List<string> errors)
        {
            errors = new List<string>();

            if (!File.Exists(filePath))
            {
                errors.Add("File not found");
                return false;
            }

            try
            {
                var lines = File.ReadAllLines(filePath);
                if (lines.Length == 0)
                {
                    errors.Add("File is empty");
                    return false;
                }

                if (!lines[0].StartsWith("Category"))
                {
                    errors.Add("Invalid header");
                }

                for (int i = 1; i < lines.Length; i++)
                {
                    var parts = lines[i].Split('\t');
                    if (parts.Length < 3)
                    {
                        errors.Add($"Line {i + 1}: Not enough columns");
                    }
                }

                return errors.Count == 0;
            }
            catch (Exception ex)
            {
                errors.Add($"Error reading file: {ex.Message}");
                return false;
            }
        }

        private object GetExportLayerTable(DWGExportOptions options)
        {
            try
            {
                var method = options.GetType().GetMethod("GetExportLayerTable");
                if (method != null)
                    return method.Invoke(options, null);

                var prop = options.GetType().GetProperty("ExportLayerTable");
                if (prop != null)
                    return prop.GetValue(options);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LayerMapping] GetExportLayerTable error: {ex.Message}");
            }
            return null;
        }

        private string ExtractLayerEntry(object entry)
        {
            try
            {
                var entryType = entry.GetType();

                string category = GetPropertyValue(entry, "Category") ?? GetPropertyValue(entry, "LayerCategory") ?? "";
                string subCategory = GetPropertyValue(entry, "SubCategory") ?? "";
                string layerName = GetPropertyValue(entry, "LayerName") ?? GetPropertyValue(entry, "Layer") ?? "";
                string colorId = GetPropertyValue(entry, "ColorId") ?? "0";
                string linetype = GetPropertyValue(entry, "Linetype") ?? GetPropertyValue(entry, "LineType") ?? "Continuous";

                if (!string.IsNullOrEmpty(layerName))
                    return $"{category}\t{subCategory}\t{layerName}\t{colorId}\t{linetype}";
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LayerMapping] Extract error: {ex.Message}");
            }
            return null;
        }

        private string GetPropertyValue(object obj, string propertyName)
        {
            try
            {
                var prop = obj.GetType().GetProperty(propertyName,
                    BindingFlags.Public | BindingFlags.Instance);
                if (prop != null)
                {
                    var val = prop.GetValue(obj);
                    return val?.ToString() ?? "";
                }
            }
            catch { }
            return null;
        }

        private object CreateLayerEntry(Type entryType, LayerMappingEntry imported)
        {
            try
            {
                var entry = Activator.CreateInstance(entryType);
                if (entry == null) return null;

                SetPropertyValue(entry, "Category", imported.Category);
                SetPropertyValue(entry, "SubCategory", imported.SubCategory);
                SetPropertyValue(entry, "LayerName", imported.LayerName);
                SetPropertyValue(entry, "ColorId", imported.ColorId);
                SetPropertyValue(entry, "Linetype", imported.Linetype);

                return entry;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[LayerMapping] Create entry error: {ex.Message}");
                return null;
            }
        }

        private void SetPropertyValue(object obj, string propertyName, object value)
        {
            try
            {
                var prop = obj.GetType().GetProperty(propertyName,
                    BindingFlags.Public | BindingFlags.Instance);
                if (prop != null && prop.CanWrite)
                {
                    if (prop.PropertyType == typeof(int) && value is int)
                        prop.SetValue(obj, value);
                    else if (prop.PropertyType == typeof(string))
                        prop.SetValue(obj, value?.ToString() ?? "");
                    else
                        prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType));
                }
            }
            catch { }
        }
    }
}
