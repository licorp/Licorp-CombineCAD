using System.Collections.Generic;

namespace Licorp_MergeSheets
{
    public class MergeConfig
    {
        public string Mode { get; set; }
        public string OutputPath { get; set; }
        public string VerticalAlign { get; set; } = "Top";
        public string DwgVersion { get; set; } = "Current";
        public int ExpectedSheetCount { get; set; }
        public bool VerifyAfterSave { get; set; } = true;
        public bool SheetSetEnabled { get; set; } = true;
        public string SheetSetIndexPath { get; set; }
        public string RasterImageMode { get; set; } = "KeepReference";
        public string ViewportMode { get; set; } = "Live";
        public string StatusPath { get; set; }
        public List<SourceFile> SourceFiles { get; set; }
    }

    public class SourceFile
    {
        public string Path { get; set; }
        public string Layout { get; set; }
    }
}
