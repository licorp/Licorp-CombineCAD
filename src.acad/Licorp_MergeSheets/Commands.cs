using System;
using System.IO;
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

            try
            {
                AcadLogger.LogSection("LICORP_MERGESHEETS Command Started");
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

                if (!File.Exists(configPath))
                {
                    AcadLogger.LogError($"Config file not found: {configPath}");
                    return;
                }

                var configJson = File.ReadAllText(configPath);
                AcadLogger.LogDebug($"Config JSON length: {configJson.Length}");

                var config = JsonConvert.DeserializeObject<MergeConfig>(configJson);

                if (config == null)
                {
                    AcadLogger.LogError("Failed to deserialize config");
                    return;
                }

                AcadLogger.LogSection("Merge Configuration");
                AcadLogger.LogInfo($"Mode: {config.Mode}");
                AcadLogger.LogInfo($"Output: {config.OutputPath}");
                AcadLogger.LogInfo($"Source files: {config.SourceFiles?.Count ?? 0}");
                AcadLogger.LogInfo($"DwgVersion: {config.DwgVersion}");

                if (config.SourceFiles != null)
                {
                    for (int i = 0; i < config.SourceFiles.Count; i++)
                    {
                        var sf = config.SourceFiles[i];
                        AcadLogger.LogDebug($" [{i}] {sf.Path} -> Layout: {sf.Layout}");
                    }
                }

                AcadLogger.LogSection("Starting Merge Operation");
                var merger = new LayoutMerger();
                bool success = false;

                switch (config.Mode)
                {
                    case "MultiLayout":
                        AcadLogger.LogInfo("Calling MergeToMultiLayout...");
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
                        return;
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
                    AcadLogger.LogError("Merge FAILED - check logs above for details");
                }
            }
            catch (System.Exception ex)
            {
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
                if (silentMode && File.Exists(SilentConfigPath))
                {
                    try { File.Delete(SilentConfigPath); }
                    catch { }
                }

                AcadLogger.LogSection("Command Finished");
            }
        }
    }
}