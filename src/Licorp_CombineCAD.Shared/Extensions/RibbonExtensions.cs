using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Licorp_CombineCAD.Services;
namespace Licorp_CombineCAD.Extensions
{
    public static class RibbonExtensions
    {
        public static PushButtonDataBuilder AddPushButton(this RibbonPanel panel, string name, string text)
        {
            return new PushButtonDataBuilder(panel, name, text);
        }
    }

    public class PushButtonDataBuilder
    {
        private readonly RibbonPanel _panel;
        private readonly string _name;
        private readonly string _text;
        private string _toolTip;
        private string _longDescription;
        private string _className;
        private string _iconName;
        private BitmapSource _largeImage;
        private BitmapSource _smallImage;

        public PushButtonDataBuilder(RibbonPanel panel, string name, string text)
        {
            _panel = panel;
            _name = name;
            _text = text;
        }

        public PushButtonDataBuilder WithCommand<T>() where T : IExternalCommand
        {
            _className = typeof(T).FullName;
            return this;
        }

        public PushButtonDataBuilder WithCommand(string className)
        {
            _className = className;
            return this;
        }

        public PushButtonDataBuilder WithToolTip(string toolTip)
        {
            _toolTip = toolTip;
            return this;
        }

        public PushButtonDataBuilder WithLongDescription(string longDescription)
        {
            _longDescription = longDescription;
            return this;
        }

        public PushButtonDataBuilder WithIcon(string iconName)
        {
            _iconName = iconName;
            return this;
        }

        public PushButtonDataBuilder WithImage(BitmapSource largeImage, BitmapSource smallImage)
        {
            _largeImage = largeImage;
            _smallImage = smallImage;
            return this;
        }

        public PushButton Build()
        {
            if (string.IsNullOrEmpty(_className))
                throw new InvalidOperationException("Command class must be specified via WithCommand<T>() or WithCommand()");

            var assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var data = new PushButtonData(_name, _text, assemblyPath, _className);

            if (!string.IsNullOrEmpty(_toolTip))
                data.ToolTip = _toolTip;

            if (!string.IsNullOrEmpty(_longDescription))
                data.LongDescription = _longDescription;

            var button = _panel.AddItem(data) as PushButton;

            if (!string.IsNullOrEmpty(_iconName))
            {
                try
                {
                    button.LargeImage = IconLoader.LoadIcon(_iconName, 32);
                    button.Image = IconLoader.LoadIcon(_iconName, 16);
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"Failed to set icon for button: {_name} - {ex.Message}");
                }
            }
            else if (_largeImage != null || _smallImage != null)
            {
                if (_largeImage != null) button.LargeImage = _largeImage;
                if (_smallImage != null) button.Image = _smallImage;
            }

            return button;
        }
    }
}
