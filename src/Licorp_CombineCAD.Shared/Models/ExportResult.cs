using System.Collections.Generic;
using System.Linq;

namespace Licorp_CombineCAD.Models
{
    public class ExportResult
    {
        public List<string> ExportedFiles { get; set; } = new List<string>();
        public List<string> FailedSheets { get; set; } = new List<string>();
        public List<string> SkippedSheets { get; set; } = new List<string>();

        public bool HasWarnings => FailedSheets.Count > 0 || SkippedSheets.Count > 0;
        public int TotalProcessed => ExportedFiles.Count + FailedSheets.Count + SkippedSheets.Count;
        public string Summary => HasWarnings
            ? $"Exported {ExportedFiles.Count}, Failed {FailedSheets.Count}, Skipped {SkippedSheets.Count}"
            : $"Exported {ExportedFiles.Count} file(s) successfully";
    }
}
