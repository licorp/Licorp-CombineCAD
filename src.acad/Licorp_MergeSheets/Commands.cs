using System;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

namespace Licorp_MergeSheets
{
    public class MergeCommands
    {
        private static readonly string SilentConfigPath = Path.Combine(Path.GetTempPath(), "Licorp_MergeSheets_Config.json");

        [CommandMethod("LICORP_MERGESHEETS", CommandFlags.Session | CommandFlags.NoHistory)]
        public void MergeSheetsCommand()
        {
            string configPath = null;
            bool silentMode = false;
            MergeConfig config = null;
            bool success = false;
            string statusMessage = null;

            try
            {
                AcadLogger.LogSection("LICORP_MERGESHEETS Command Started");
                AcadLogger.LogInfo("[ACAD-CMD] command=LICORP_MERGESHEETS");
                AcadLogger.LogInfo($"Log file: {AcadLogger.GetLogFilePath()}");

                silentMode = File.Exists(SilentConfigPath);

                if (!silentMode)
                {
                    var doc = Application.DocumentManager.MdiActiveDocument;
                    if (doc == null)
                    {
                        AcadLogger.LogError("No active document");
                        return;
                    }

                    var ed = doc.Editor;
                    var pr = ed.GetString("Enter config file path: ");
                    if (pr.Status != PromptStatus.OK) return;
                    configPath = pr.StringResult;
                }
                else
                {
                    configPath = SilentConfigPath;
                    AcadLogger.LogInfo("Silent mode: reading config from temp file");
                }

                AcadLogger.LogInfo($"[ACAD-CMD] silentMode={silentMode}, configPath={configPath}");

                if (!File.Exists(configPath))
                {
                    AcadLogger.LogError($"Config file not found: {configPath}");
                    return;
                }

                var configJson = File.ReadAllText(configPath);
                AcadLogger.LogDebug($"Config JSON length: {configJson.Length}");

                config = JsonConvert.DeserializeObject<MergeConfig>(configJson);

                if (config == null)
                {
                    AcadLogger.LogError("Failed to deserialize config");
                    statusMessage = "Failed to deserialize merge config.";
                    return;
                }

                AcadLogger.LogSection("Merge Configuration");
                AcadLogger.LogInfo($"[ACAD-CMD] statusPath={config.StatusPath}");
                AcadLogger.LogInfo($"Mode: {config.Mode}");
                AcadLogger.LogInfo($"Output: {config.OutputPath}");
                AcadLogger.LogInfo($"Source files: {config.SourceFiles?.Count ?? 0}");
                
                // Warning for large batch processing
                int sheetCount = config.SourceFiles?.Count ?? 0;
                if (sheetCount > 50)
                {
                    AcadLogger.LogWarning($"LARGE BATCH: Processing {sheetCount} sheets. This may take {sheetCount * 2}+ minutes.");
                    AcadLogger.LogWarning($"Ensure sufficient disk space and do not interrupt the process.");
                }
                
                AcadLogger.LogInfo($"DwgVersion: {config.DwgVersion}");
                AcadLogger.LogInfo($"ExpectedSheetCount: {config.ExpectedSheetCount}");
                AcadLogger.LogInfo($"VerifyAfterSave: {config.VerifyAfterSave}");
                AcadLogger.LogInfo($"CombinedDwgIndexEnabled: {config.SheetSetEnabled}");
                AcadLogger.LogInfo($"RasterImageMode: {config.RasterImageMode}");
                AcadLogger.LogInfo($"MergeLayers: {config.MergeLayers}");
                AcadLogger.LogInfo($"LayoutNamingRule: {config.LayoutNamingRule}");
                if (!string.IsNullOrWhiteSpace(config.LayoutNamingPattern))
                    AcadLogger.LogInfo($"LayoutNamingPattern: {config.LayoutNamingPattern}");
                if (!string.IsNullOrWhiteSpace(config.LayoutNamingPrefix))
                    AcadLogger.LogInfo($"LayoutNamingPrefix: {config.LayoutNamingPrefix}");
                AcadLogger.LogInfo($"ViewportMode: {config.ViewportMode}");

                if (!string.IsNullOrWhiteSpace(config.SourceFolder))
                    AcadLogger.LogInfo($"SourceFolder: {config.SourceFolder}");
                if (!string.IsNullOrWhiteSpace(config.SourcePattern))
                    AcadLogger.LogInfo($"SourcePattern: {config.SourcePattern}");
                if (config.BackupBeforeOverwrite)
                    AcadLogger.LogInfo("BackupBeforeOverwrite: enabled");
                if (config.TitleBlockAutoFill)
                    AcadLogger.LogInfo($"TitleBlockAutoFill: enabled (CSV={config.TitleBlockCsvPath})");
                if (config.ApplyLayerMapping)
                    AcadLogger.LogInfo($"ApplyLayerMapping: enabled ({config.LayerMappingRules?.Count ?? 0} rules)");
                if (!string.IsNullOrWhiteSpace(config.LayoutNamingPreset))
                    AcadLogger.LogInfo($"LayoutNamingPreset: {config.LayoutNamingPreset}");
                if (config.AutoPdfExport)
                    AcadLogger.LogInfo($"AutoPdfExport: enabled (folder={config.PdfOutputFolder})");

                if (!string.IsNullOrWhiteSpace(config.LayoutNamingPreset))
                {
                    string presetPattern = LayoutNamingPresets.GetPattern(config.LayoutNamingPreset);
                    if (presetPattern != null)
                    {
                        config.LayoutNamingRule = "Custom";
                        config.LayoutNamingPattern = presetPattern;
                        AcadLogger.LogInfo($"Applied preset '{config.LayoutNamingPreset}' -> pattern '{presetPattern}'");
                    }
                    else
                    {
                        AcadLogger.LogWarning($"Unknown layout naming preset: {config.LayoutNamingPreset}");
                    }
                }

                if (!string.IsNullOrWhiteSpace(config.SourceFolder) &&
                    (config.SourceFiles == null || config.SourceFiles.Count == 0))
                {
                    AcadLogger.LogSection("Batch Folder Scan");
                    var folderService = new BatchFolderService();
                    config.SourceFiles = folderService.ScanFolder(
                        config.SourceFolder,
                        config.SourcePattern ?? "*.dwg",
                        config.RecursiveScan);
                    AcadLogger.LogInfo($"Batch scan found {config.SourceFiles.Count} file(s)");
                }

                if (config.PreflightCheck)
                {
                    AcadLogger.LogSection("Preflight Check");
                    var preflightService = new PreflightService();
                    var preflightResult = preflightService.RunPreflight(config);

                    foreach (var warning in preflightResult.Warnings)
                        AcadLogger.LogWarning($"Preflight: {warning}");

                    if (!preflightResult.Success)
                    {
                        foreach (var error in preflightResult.Errors)
                            AcadLogger.LogError($"Preflight: {error}");
                        statusMessage = "Preflight check failed: " + string.Join("; ", preflightResult.Errors);
                        return;
                    }

                    if (config.ExpectedSheetCount > 0 && preflightResult.ValidFileCount != config.ExpectedSheetCount)
                    {
                        var lockedWarnings = preflightResult.Warnings
                            .Where(w => w.IndexOf("locked", StringComparison.OrdinalIgnoreCase) >= 0)
                            .ToList();

                        var msg = $"Preflight failed: expected {config.ExpectedSheetCount} source DWG(s), " +
                                  $"but only {preflightResult.ValidFileCount} valid. " +
                                  (lockedWarnings.Count > 0
                                      ? string.Join(" | ", lockedWarnings)
                                      : $"{config.ExpectedSheetCount - preflightResult.ValidFileCount} file(s) missing or invalid.");

                        AcadLogger.LogError(msg);
                        statusMessage = msg;
                        success = false;
                        return;
                    }
                }

                if (config.SourceFiles != null)
                {
                    for (int i = 0; i < config.SourceFiles.Count; i++)
                    {
                        var sf = config.SourceFiles[i];
                        AcadLogger.LogDebug($" [{i}] {sf.Path} -> Layout: {sf.Layout}");
                    }
                }

                if (config.BackupBeforeOverwrite && File.Exists(config.OutputPath))
                {
                    AcadLogger.LogSection("Backup");
                    var backupService = new BackupService();
                    string backupPath = backupService.CreateBackup(config.OutputPath);
                    if (backupPath != null)
                        AcadLogger.LogInfo($"Backup created: {backupPath}");
                }

                AcadLogger.LogSection("Starting Merge Operation");
                var merger = new LayoutMerger();

                switch (config.Mode)
                {
                    case "MultiLayout":
                        AcadLogger.LogInfo("Calling MergeToMultiLayout...");
                        success = merger.MergeToMultiLayout(config);
                        break;
                    case "ModelFirstMultiLayout":
                        AcadLogger.LogInfo("Calling MergeToMultiLayout (ModelFirstMultiLayout alias)...");
                        success = merger.MergeToMultiLayout(config);
                        break;
                    case "SingleLayout":
                        AcadLogger.LogInfo("Calling MergeToSingleLayout...");
                        success = merger.MergeToSingleLayout(config);
                        break;
                    case "ModelSpace":
                        AcadLogger.LogInfo("Calling MergeToModelSpace...");
                        success = merger.MergeToModelSpace(config);
                        break;
                    default:
                        AcadLogger.LogError($"Unknown mode: {config.Mode}");
                        statusMessage = $"Unknown merge mode: {config.Mode}";
                        return;
                }

                AcadLogger.LogInfo($"[ACAD-CMD] mergeResult success={success}, mode={config.Mode}");

                if (success && config.VerifyAfterSave)
                {
                    AcadLogger.LogSection("Post-Save Verification");
                    success = merger.VerifyCombinedFile(config, out statusMessage);
                }

                if (success)
                {
                    merger.HandleRasterImages(config);
                    merger.CreateCombinedDwgIndex(config);

                    if (config.TitleBlockAutoFill && !string.IsNullOrWhiteSpace(config.TitleBlockCsvPath))
                    {
                        AcadLogger.LogSection("Title Block Auto-Fill");
                        var titleBlockService = new TitleBlockFillService();
                        var mappings = titleBlockService.LoadMappingsFromCsv(config.TitleBlockCsvPath, config.TitleBlockAttributeName);
                        if (mappings.Count > 0)
                        {
                            int filled = titleBlockService.FillTitleBlocks(config.OutputPath, mappings);
                            AcadLogger.LogInfo($"Title block auto-fill: {filled} attribute(s) filled");
                        }
                    }

                    if (config.AutoPdfExport)
                    {
                        AcadLogger.LogSection("Auto PDF Export");
                        var pdfService = new PdfExportService();
                        string pdfFolder = config.PdfOutputFolder ?? Path.GetDirectoryName(config.OutputPath);
                        pdfService.SchedulePdfExport(config.OutputPath, pdfFolder, config.PdfPresetName);
                    }
                }

                if (success)
                {
                    AcadLogger.LogSection("Merge Completed Successfully");
                    AcadLogger.LogInfo($"Output file: {config.OutputPath}");

                    if (File.Exists(config.OutputPath))
                    {
                        var fileInfo = new FileInfo(config.OutputPath);
                        AcadLogger.LogInfo($"File size: {fileInfo.Length / 1024.0:F2} KB");
                    }
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(statusMessage))
                        statusMessage = merger.LastError ?? "Merge failed. Check merge log for details.";
                    AcadLogger.LogError($"Merge FAILED: {statusMessage}");
                }
            }
            catch (System.Exception ex)
            {
                success = false;
                statusMessage = ex.Message;
                if (ex.InnerException != null)
                    statusMessage += " | Inner: " + ex.InnerException.Message;
                AcadLogger.LogSection("EXCEPTION CAUGHT");
                AcadLogger.LogError($"Message: {ex.Message}");
                AcadLogger.LogError($"Type: {ex.GetType().FullName}");
                AcadLogger.LogError($"Stack Trace:\n{ex.StackTrace}");

                if (ex.InnerException != null)
                {
                    AcadLogger.LogError($"Inner Exception: {ex.InnerException.Message}");
                }
            }
            finally
            {
                WriteStatus(config, success, statusMessage);
                AcadLogger.LogInfo($"[ACAD-CMD] finalStatus success={success}, message={statusMessage}");

                if (silentMode && File.Exists(SilentConfigPath))
                {
                    try { File.Delete(SilentConfigPath); }
                    catch { }
                }

                AcadLogger.LogSection("Command Finished");
            }
        }

        private void WriteStatus(MergeConfig config, bool success, string message)
        {
            try
            {
                if (config == null)
                {
                    AcadLogger.LogWarning("WriteStatus skipped: config is null.");
                    return;
                }

                var effectiveStatusPath = config.StatusPath;
                if (string.IsNullOrWhiteSpace(effectiveStatusPath))
                {
                    effectiveStatusPath = Path.Combine(Path.GetTempPath(), "Licorp_MergeSheets_Status.json");
                    AcadLogger.LogWarning($"StatusPath was empty, using fallback: {effectiveStatusPath}");
                }

                var finalMessage = string.IsNullOrWhiteSpace(message)
                    ? (success ? "Merge completed successfully." : "Merge failed.")
                    : message;

                var logPath = AcadLogger.GetLogFilePath();

                var status = new
                {
                    success = success,
                    output = config.OutputPath,
                    error = success ? null : finalMessage,
                    log = logPath,
                    timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),

                    Success = success,
                    Message = finalMessage,
                    OutputPath = config.OutputPath,
                    LogPath = logPath
                };

                var json = JsonConvert.SerializeObject(status, Formatting.Indented);

                var tempPath = effectiveStatusPath + ".tmp";

                Directory.CreateDirectory(Path.GetDirectoryName(effectiveStatusPath));

                File.WriteAllText(tempPath, json);

                if (File.Exists(effectiveStatusPath))
                    File.Delete(effectiveStatusPath);

                File.Move(tempPath, effectiveStatusPath);

                using (var fs = new FileStream(effectiveStatusPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    fs.Flush(true);
                }

                AcadLogger.LogInfo($"Status written: {effectiveStatusPath} (success={success})");
                AcadLogger.LogDebug($"Status JSON:\n{json}");

                System.Threading.Thread.Sleep(200);
            }
            catch (System.Exception ex)
            {
                AcadLogger.LogError($"CRITICAL: Failed to write status file: {ex.Message}");
                AcadLogger.LogError($"Status path: {config?.StatusPath}");

                try
                {
                    var emergencyPath = Path.Combine(Path.GetTempPath(), "Licorp_EmergencyStatus.json");
                    File.WriteAllText(emergencyPath, JsonConvert.SerializeObject(new
                    {
                        success = false,
                        error = $"Status write failed: {ex.Message}",
                        originalStatusPath = config?.StatusPath,
                        timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),

                        Success = false,
                        Message = $"Status write failed: {ex.Message}",
                        OriginalStatusPath = config?.StatusPath,
                        Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    }, Formatting.Indented));
                    AcadLogger.LogInfo($"Emergency status written to: {emergencyPath}");
                }
                catch (System.Exception emergencyEx)
                {
                    AcadLogger.LogError($"Emergency status write also failed: {emergencyEx.Message}");
                }
            }
        }
    }
}
