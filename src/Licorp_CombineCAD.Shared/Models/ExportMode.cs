namespace Licorp_CombineCAD.Models
{
    /// <summary>
    /// Export modes supported by CombineCAD
    /// </summary>
    public enum ExportMode
    {
        /// <summary>
        /// Export each sheet as individual DWG file
        /// </summary>
        Individual,

        /// <summary>
        /// Export sheets to individual DWGs then merge into 1 file with multiple layouts
        /// (each sheet = 1 layout in AutoCAD)
        /// Requires AutoCAD (AcCoreConsole.exe)
        /// </summary>
        MultiLayout,

        /// <summary>
        /// Combine all sheets into 1 DWG file with 1 layout
        /// (sheets arranged side-by-side)
        /// Requires AutoCAD (AcCoreConsole.exe)
        /// </summary>
        SingleLayout,

        /// <summary>
        /// Export all sheets into Model Space of a single DWG
        /// (with title blocks, arranged in grid layout)
        /// Requires AutoCAD (AcCoreConsole.exe)
        /// </summary>
        ModelSpace
    }
}
