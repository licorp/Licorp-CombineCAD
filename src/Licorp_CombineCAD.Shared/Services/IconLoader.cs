using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
namespace Licorp_CombineCAD.Services
{
    public static class IconLoader
    {
        private static readonly Dictionary<string, BitmapSource> _iconCache = new Dictionary<string, BitmapSource>();

        public static BitmapSource LoadIcon(string iconName, int size)
        {
            var cacheKey = $"{iconName}_{size}";
            if (_iconCache.TryGetValue(cacheKey, out var cached) && cached != null)
                return cached;

            try
            {
                var uri = new Uri(
                    $"pack://application:,,,/Licorp_CombineCAD;component/Resources/Icons/{iconName}_{size}.png",
                    UriKind.Absolute);

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = uri;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();

                _iconCache[cacheKey] = bitmap;
                Logger.LogDebug($"Icon loaded via Pack URI: {iconName} ({size}px)");
                return bitmap;
            }
            catch
            {
                var fallback = GeneratePlaceholderIcon(iconName, size);
                if (fallback != null)
                {
                    _iconCache[cacheKey] = fallback;
                    Logger.LogDebug($"Icon generated as fallback: {iconName} ({size}px)");
                }
                return fallback;
            }
        }

        private static BitmapSource GeneratePlaceholderIcon(string iconName, int size)
        {
            try
            {
                var bgColor = GetIconColor(iconName);

                var brush = new SolidColorBrush(bgColor);
                var visual = new DrawingVisual();
                using (var dc = visual.RenderOpen())
                {
                    dc.DrawRectangle(brush, null, new Rect(0, 0, size, size));

                    var typeface = new Typeface("Segoe UI");
                    var fontSize = size <= 16 ? 6 : 12;
                    var text = GetIconText(iconName);
                    var formatted = new FormattedText(
                        text,
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
                Logger.LogWarning($"Placeholder icon generation failed: {iconName} - {ex.Message}");
                return null;
            }
        }

        private static Color GetIconColor(string iconName)
        {
            return iconName switch
            {
                "multi_layout" => Colors.DodgerBlue,
                "single_layout" => Colors.MediumOrchid,
                "model_space" => Colors.DarkOrange,
                "layers" => Colors.Gold,
                _ => Colors.Gray
            };
        }

        private static string GetIconText(string iconName)
        {
            return iconName switch
            {
                "multi_layout" => "ML",
                "single_layout" => "SL",
                "model_space" => "MS",
                "layers" => "LY",
                _ => iconName.Substring(0, Math.Min(2, iconName.Length)).ToUpper()
            };
        }
    }
}
