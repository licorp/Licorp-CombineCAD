using System;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.Services;
using Licorp_CombineCAD.Views;

namespace Licorp_CombineCAD.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ExportMultiLayoutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExportCommandBase.Execute(commandData, ExportMode.MultiLayout, ref message);
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ExportSingleLayoutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExportCommandBase.Execute(commandData, ExportMode.SingleLayout, ref message);
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ExportModelSpaceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExportCommandBase.Execute(commandData, ExportMode.ModelSpace, ref message);
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class LayerManagerCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Logger.LogSection("LayerManager Command");
                var uiDoc = commandData.Application.ActiveUIDocument;
                if (uiDoc == null)
                {
                    TaskDialog.Show("CombineCAD", "Please open a document first.");
                    return Result.Cancelled;
                }

                var dialog = new LayerManagerDialog(uiDoc);
                dialog.Show();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Logger.LogException(ex, "LayerManager");
                return Result.Failed;
            }
        }
    }

    internal static class ExportCommandBase
    {
        public static Result Execute(ExternalCommandData commandData, ExportMode mode, ref string message)
        {
            try
            {
                Logger.LogSection($"Export Command: {mode}");
                Logger.LogInfo($"User: {Environment.UserName}");
                Logger.LogInfo($"Machine: {Environment.MachineName}");

                var uiDoc = commandData.Application.ActiveUIDocument;
                if (uiDoc == null)
                {
                    TaskDialog.Show("CombineCAD", "Please open a document first.");
                    return Result.Cancelled;
                }

                var doc = uiDoc.Document;
                Logger.LogInfo($"Document: {doc.Title}");
                Logger.LogInfo($"Worksharing: {doc.IsWorkshared}");

                var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .WhereElementIsNotElementType();

                int sheetCount = collector.GetElementCount();
                Logger.LogInfo($"Found {sheetCount} sheets");

                if (sheetCount == 0)
                {
                    TaskDialog.Show("CombineCAD", "No sheets found in the current document.");
                    return Result.Cancelled;
                }

                var dialog = new ExportDialog(uiDoc, mode);
                dialog.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Logger.LogException(ex, "ExportCommand");
                return Result.Failed;
            }
        }
    }
}