namespace Licorp_CombineCAD.Models
{
    /// <summary>
    /// Export modes supported by CombineCAD
    /// </summary>
    public enum ExportMode
    {
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
        ModelSpace,

        /// <summary>
        /// Export sheets with exact 1:1 scale - no scaling applied.
        /// Paper size = Revit sheet size, viewport custom scale = 1.0.
        /// When plotting: use 1:1 scale on paper.
        /// </summary>
        OneToOneScale,

        /// <summary>
        /// Export sheets with sheet-ratio-based scaling.
        /// Calculates ratio between sheet size and primary viewport size,
        /// then applies that ratio to fit content into paper.
        /// Useful when viewport does not fit the original sheet.
        /// </summary>
        SheetRatioScale
    }
}
