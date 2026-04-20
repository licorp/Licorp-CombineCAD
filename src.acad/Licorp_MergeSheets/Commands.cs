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
        [CommandMethod("LICORP_MERGESHEETS")]
        public void MergeSheetsCommand(string configPath = null)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[Licorp_MergeSheets] Command started");

                if (string.IsNullOrEmpty(configPath))
                {
                    var doc = Application.DocumentManager.MdiActiveDocument;
                    if (doc == null) return;

                    var ed = doc.Editor;
                    var pr = ed.GetString("Enter config file path: ");
                    if (pr.Status != PromptStatus.OK) return;
                    configPath = pr.StringResult;
                }

                if (!File.Exists(configPath))
                {
                    System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Config not found: {configPath}");
                    return;
                }

                var configJson = File.ReadAllText(configPath);
                var config = JsonConvert.DeserializeObject<MergeConfig>(configJson);

                if (config == null)
                {
                    System.Diagnostics.Debug.WriteLine("[Licorp_MergeSheets] Invalid config");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Mode: {config.Mode}");
                System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Source files: {config.SourceFiles.Count}");
                System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Output: {config.OutputPath}");

                var merger = new LayoutMerger();
                bool success = false;

                switch (config.Mode)
                {
                    case "MultiLayout":
                        success = merger.MergeToMultiLayout(config);
                        break;
                    case "SingleLayout":
                        success = merger.MergeToSingleLayout(config);
                        break;
                    case "ModelSpace":
                        success = merger.MergeToModelSpace(config);
                        break;
                }

                System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Merge result: {success}");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[Licorp_MergeSheets] Stack: {ex.StackTrace}");
            }
        }
    }
}