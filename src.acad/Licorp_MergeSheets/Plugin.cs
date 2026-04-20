using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(Licorp_MergeSheets.Plugin))]
[assembly: CommandClass(typeof(Licorp_MergeSheets.MergeCommands))]

namespace Licorp_MergeSheets
{
    public class Plugin : IExtensionApplication
    {
        public void Initialize()
        {
            System.Diagnostics.Debug.WriteLine("[Licorp_MergeSheets] Plugin loaded");
        }

        public void Terminate()
        {
        }
    }
}