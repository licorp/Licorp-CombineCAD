using System.Collections.Generic;

namespace Licorp_MergeSheets
{
    public class MergeConfig
    {
        public string Mode { get; set; }
        public string OutputPath { get; set; }
        public string VerticalAlign { get; set; } = "Top";
        public List<SourceFile> SourceFiles { get; set; }
    }

    public class SourceFile
    {
        public string Path { get; set; }
        public string Layout { get; set; }
    }
}