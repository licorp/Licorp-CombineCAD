using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Licorp_CombineCAD.Services;

namespace Licorp_CombineCAD
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            Logger.Initialize();
            Logger.LogSection("LICORP_COMBINECAD Add-in Starting");
            Logger.LogInfo($"Revit Version: {application.ControlledApplication.VersionNumber}");
            Logger.LogInfo($"UI Culture: {System.Globalization.CultureInfo.CurrentUICulture.Name}");

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            if (string.IsNullOrEmpty(assemblyPath))
            {
#pragma warning disable SYSLIB0012 // Assembly.CodeBase is obsolete but needed for .NET Framework fallback
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
#pragma warning restore SYSLIB0012
                if (!string.IsNullOrEmpty(codeBase) && codeBase.StartsWith("file:///"))
                {
                    assemblyPath = new Uri(codeBase).LocalPath;
                }
                else if (!string.IsNullOrEmpty(codeBase))
                {
                    assemblyPath = codeBase.Replace("file://", "").TrimStart('/');
                    if (assemblyPath.Length > 2 && assemblyPath[1] != ':')
                        assemblyPath = @"C:\" + assemblyPath;
                }
            }

            Logger.LogInfo($"Assembly path: {assemblyPath}");

            string tabName = "Licorp";
            try
            {
                application.CreateRibbonTab(tabName);
                Logger.LogInfo($"Created ribbon tab: {tabName}");
            }
            catch
            {
                Logger.LogWarning($"Ribbon tab already exists: {tabName}");
            }

            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Combine CAD");

            string ns = "Licorp_CombineCAD.Commands";

            var multiLayoutData = new PushButtonData(
                "ExportMultiLayout",
                "Multi-Layout\nDWG",
                assemblyPath,
                $"{ns}.ExportMultiLayoutCommand");
            multiLayoutData.ToolTip = "Export sheets to 1 DWG file with multiple layouts";
            multiLayoutData.LongDescription = "Export selected Revit sheets to individual DWG files, " +
                "then automatically merge them into a single DWG with multiple layouts " +
                "(each sheet = 1 layout). Requires AutoCAD.";
            var multiBtn = panel.AddItem(multiLayoutData) as PushButton;
            SetButtonIcon(multiBtn, "multi_layout", Colors.DodgerBlue);

            var singleLayoutData = new PushButtonData(
                "ExportSingleLayout",
                "Single Layout\nDWG",
                assemblyPath,
                $"{ns}.ExportSingleLayoutCommand");
            singleLayoutData.ToolTip = "Combine all sheets into 1 DWG with 1 layout";
            var singleBtn = panel.AddItem(singleLayoutData) as PushButton;
            SetButtonIcon(singleBtn, "single_layout", Colors.MediumOrchid);

            var modelSpaceData = new PushButtonData(
                "ExportModelSpace",
                "Model Space\nDWG",
                assemblyPath,
                $"{ns}.ExportModelSpaceCommand");
            modelSpaceData.ToolTip = "Export sheets to Model Space with title blocks";
            var msBtn = panel.AddItem(modelSpaceData) as PushButton;
            SetButtonIcon(msBtn, "model_space", Colors.DarkOrange);

            panel.AddSeparator();

            var layerData = new PushButtonData(
                "LayerManager",
                "Layer\nManager",
                assemblyPath,
                $"{ns}.LayerManagerCommand");
            layerData.ToolTip = "Export/Import DWG Export Layer settings";
            layerData.LongDescription = "Save and load DWG export layer mapping to/from .txt files " +
                "for sharing across projects and team members.";
            var layerBtn = panel.AddItem(layerData) as PushButton;
            SetButtonIcon(layerBtn, "layers", Colors.Gold);

            Logger.LogInfo("Ribbon setup completed");
            Logger.LogSection("Add-in Ready");
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Logger.LogInfo("Add-in shutting down");
            return Result.Succeeded;
        }

        private void SetButtonIcon(RibbonButton button, string iconName, Color fallbackColor)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string resourceName = $"Licorp_CombineCAD.Resources.Icons.{iconName}_32.png";

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = stream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        button.LargeImage = bitmap;
                        Trace.WriteLine($"[CombineCAD] Loaded icon: {resourceName}");
                    }
                    else
                    {
                        button.LargeImage = CreateTextIcon(iconName.Replace('_', ' ').ToUpper(), fallbackColor, 32);
                    }
                }

                string smallResource = $"Licorp_CombineCAD.Resources.Icons.{iconName}_16.png";
                using (var stream = assembly.GetManifestResourceStream(smallResource))
                {
                    if (stream != null)
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = stream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        button.Image = bitmap;
                    }
                    else
                    {
                        button.Image = CreateTextIcon(iconName.Replace('_', ' ').ToUpper(), fallbackColor, 16);
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Error loading icon '{iconName}': {ex.Message}");
            }
        }

        private BitmapImage CreateTextIcon(string text, Color bgColor, int size)
        {
            try
            {
                var brush = new SolidColorBrush(bgColor);
                var visual = new DrawingVisual();
                using (var dc = visual.RenderOpen())
                {
                    dc.DrawRectangle(brush, null, new Rect(0, 0, size, size));

                    var typeface = new Typeface("Segoe UI");
                    var fontSize = size <= 16 ? 6 : 12;
                    var formatted = new FormattedText(
                        text.Length > 4 ? text.Substring(0, 4) : text,
                        System.Globalization.CultureInfo.CurrentCulture,
                        FlowDirection.LeftToRight,
                        typeface,
                        fontSize,
                        Brushes.White,
                        1.0);

                    formatted.TextAlignment = TextAlignment.Center;
                    double x = (size - formatted.Width) / 2;
                    double y = (size - formatted.Height) / 2;
                    dc.DrawText(formatted, new Point(x, y));
                }

                var renderBitmap = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
                renderBitmap.Render(visual);

                var bitmap = new BitmapImage();
                using (var stream = new MemoryStream())
                {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderBitmap));
                    encoder.Save(stream);
                    stream.Position = 0;

                    bitmap.BeginInit();
                    bitmap.StreamSource = stream;
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze();
                }

                return bitmap;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[CombineCAD] Fallback icon error: {ex.Message}");
                return null;
            }
        }
    }
}
