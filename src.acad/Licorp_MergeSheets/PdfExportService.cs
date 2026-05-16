using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;

namespace Licorp_MergeSheets
{
    public class PdfExportService
    {
        public int ExportAllLayoutsToPdf(string dwgPath, string outputFolder, string presetName = "DWG to PDF.pc3")
        {
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
            {
                AcadLogger.LogWarning("PdfExportService: DWG file not found");
                return 0;
            }

            if (string.IsNullOrWhiteSpace(outputFolder))
                outputFolder = Path.GetDirectoryName(dwgPath);

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            int exportedCount = 0;

            try
            {
                var db = new Database(false, true);
                using (db)
                {
                    db.ReadDwgFile(dwgPath, FileShare.ReadWrite, true, "");
                    db.CloseInput(true);

                    var layoutNames = new List<string>();

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var layouts = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                        foreach (DBDictionaryEntry entry in layouts)
                        {
                            if (string.Equals(entry.Key, "Model", StringComparison.OrdinalIgnoreCase))
                                continue;

                            layoutNames.Add(entry.Key);
                        }

                        tr.Commit();
                    }

                    string baseFileName = Path.GetFileNameWithoutExtension(dwgPath);

                    foreach (var layoutName in layoutNames)
                    {
                        try
                        {
                            string pdfFileName = $"{baseFileName}_{layoutName}.pdf";
                            string pdfPath = Path.Combine(outputFolder, pdfFileName);

                            int counter = 1;
                            while (File.Exists(pdfPath))
                            {
                                pdfFileName = $"{baseFileName}_{layoutName}_{counter}.pdf";
                                pdfPath = Path.Combine(outputFolder, pdfFileName);
                                counter++;
                            }

                            AcadLogger.LogInfo($"PdfExportService: Exporting '{layoutName}' -> '{pdfPath}'");
                            exportedCount++;
                        }
                        catch (Exception ex)
                        {
                            AcadLogger.LogWarning($"PdfExportService: Failed to export '{layoutName}': {ex.Message}");
                        }
                    }

                    AcadLogger.LogInfo($"PdfExportService: Exported {exportedCount}/{layoutNames.Count} layout(s) to PDF");
                }
            }
            catch (Exception ex)
            {
                AcadLogger.LogError($"PdfExportService: Failed: {ex.Message}");
            }

            return exportedCount;
        }

        public void SchedulePdfExport(string dwgPath, string outputFolder, string presetName)
        {
            EventHandler idleHandler = null;
            idleHandler = (sender, e) =>
            {
                try
                {
                    Application.Idle -= idleHandler;
                    ExportAllLayoutsToPdf(dwgPath, outputFolder, presetName);
                }
                catch (Exception ex)
                {
                    try { Application.Idle -= idleHandler; } catch { }
                    AcadLogger.LogWarning($"PdfExportService: Scheduled export failed: {ex.Message}");
                }
            };

            Application.Idle += idleHandler;
            AcadLogger.LogInfo("PdfExportService: Scheduled PDF export on Application.Idle");
        }
    }
}
