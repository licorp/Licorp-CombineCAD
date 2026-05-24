using Autodesk.Revit.UI;
using System;
using Licorp_CombineCAD.Commands;
using Licorp_CombineCAD.Extensions;
using Licorp_CombineCAD.Services;

namespace Licorp_CombineCAD
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            System.Diagnostics.Trace.WriteLine("[LicorpCAD] OnStartup called!");
            System.Diagnostics.Debug.WriteLine("[LicorpCAD] OnStartup called!");

            try
            {
                Logger.Initialize();
                Logger.LogInfo("LICORP_COMBINECAD Add-in Starting");
                Logger.LogInfo($"Revit Version: {application.ControlledApplication.VersionNumber}");
                Logger.LogInfo($"UI Culture: {System.Globalization.CultureInfo.CurrentUICulture.Name}");

                string tabName = "Licorp";
                try
                {
                    application.CreateRibbonTab(tabName);
                    Logger.LogInfo($"Created ribbon tab: {tabName}");
                }
                catch (Exception tabEx)
                {
                    Logger.LogWarning($"Ribbon tab creation failed: {tabEx.Message}");
                    ShowErrorDialog("Ribbon Tab Error", $"Failed to create tab '{tabName}': {tabEx.Message}");
                }

                var panel = application.CreateRibbonPanel(tabName, "Combine CAD");

                const string ns = "Licorp_CombineCAD.Commands";

                panel.AddPushButton("ExportMultiLayout", "Multi-Layout\nDWG")
                    .WithCommand($"{ns}.ExportMultiLayoutCommand")
                    .WithToolTip("Export sheets to 1 DWG file with multiple layouts")
                    .WithLongDescription("Export selected Revit sheets to individual DWG files, " +
                        "then automatically merge them into a single DWG with multiple layouts " +
                        "(each sheet = 1 layout). Requires AutoCAD.")
                    .WithIcon("multi_layout")
                    .Build();

                panel.AddPushButton("ExportSingleLayout", "Single Layout\nDWG")
                    .WithCommand($"{ns}.ExportSingleLayoutCommand")
                    .WithToolTip("Combine all sheets into 1 DWG with 1 layout")
                    .WithIcon("single_layout")
                    .Build();

                panel.AddPushButton("ExportModelSpace", "Model Space\nDWG")
                    .WithCommand($"{ns}.ExportModelSpaceCommand")
                    .WithToolTip("Export sheets to Model Space with title blocks")
                    .WithIcon("model_space")
                    .Build();

                panel.AddSeparator();

                panel.AddPushButton("LayerManager", "Layer\nManager")
                    .WithCommand($"{ns}.LayerManagerCommand")
                    .WithToolTip("Export/Import DWG Export Layer settings")
                    .WithLongDescription("Save and load DWG export layer mapping to/from .txt files " +
                        "for sharing across projects and team members.")
                    .WithIcon("layers")
                    .Build();

                Logger.LogInfo("Ribbon setup completed");
                Logger.LogInfo("Add-in Ready");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Logger.LogError($"OnStartup Critical Failure: {ex}");
                ShowErrorDialog("Licorp CombineCAD Load Failed", ex.ToString());
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Logger.LogInfo("Add-in shutting down");
            return Result.Succeeded;
        }

        private static void ShowErrorDialog(string title, string message)
        {
            try
            {
                var td = new TaskDialog(title)
                {
                    MainInstruction = title,
                    MainContent = message,
                    CommonButtons = TaskDialogCommonButtons.Ok,
                    AllowCancellation = true
                };
                td.Show();
            }
            catch { }
        }
    }
}
