using System;
using System.Collections.Generic;

namespace Licorp_MergeSheets
{
    public enum ArrangementMode
    {
        Horizontal,
        Vertical,
        Grid
    }

    public enum SortMode
    {
        Manual,
        SheetNumber,
        SheetName,
        Discipline,
        PaperSize,
        Reverse
    }

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
        public bool MergeLayers { get; set; } = true;
        public string LayoutNamingRule { get; set; } = "Original";
        public string LayoutNamingPattern { get; set; }
        public string LayoutNamingPrefix { get; set; }
        public string ViewportMode { get; set; } = "Live";
        public string StatusPath { get; set; }
        public List<SourceFile> SourceFiles { get; set; }

        public string SourceFolder { get; set; }
        public string SourcePattern { get; set; } = "*.dwg";
        public bool RecursiveScan { get; set; } = false;

        public bool BackupBeforeOverwrite { get; set; } = true;

        public bool TitleBlockAutoFill { get; set; } = false;
        public string TitleBlockCsvPath { get; set; }
        public string TitleBlockAttributeName { get; set; } = "SHEET_NUMBER";

        public List<LayerMappingRule> LayerMappingRules { get; set; }
        public bool ApplyLayerMapping { get; set; } = false;

        public string LayoutNamingPreset { get; set; }

        public bool AutoPdfExport { get; set; } = false;
        public string PdfOutputFolder { get; set; }
        public string PdfPresetName { get; set; } = "DWG to PDF.pc3";

        public bool PreflightCheck { get; set; } = true;

        public string TemplateDwgPath { get; set; }

        public ArrangementMode ModelSpaceArrangement { get; set; } = ArrangementMode.Horizontal;
        public int GridColumns { get; set; } = 3;
        public double CustomSpacing { get; set; } = 50.0;

        public SortMode SheetSortMode { get; set; } = SortMode.Manual;
        public bool ReverseSortOrder { get; set; } = false;

        /// <summary>
        /// Progress callback: (currentLayoutIndex, totalLayouts, layoutName)
        /// </summary>
        [NonSerialized]
        public Action<int, int, string> ProgressCallback;
    }

    public class SourceFile
    {
        public string Path { get; set; }
        public string Layout { get; set; }
        public string PaperSize { get; set; }
    }

    public class LayerMappingRule
    {
        public string SourcePattern { get; set; }
        public string TargetLayer { get; set; }
        public bool IsWildcard { get; set; } = false;
    }

    public class TitleBlockFieldMapping
    {
        public string LayoutName { get; set; }
        public string BlockName { get; set; }
        public string AttributeTag { get; set; }
        public string Value { get; set; }
    }

    public class PreflightResult
    {
        public bool Success { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public int ValidFileCount { get; set; }
        public int TotalFileCount { get; set; }
    }

    public static class LayoutNamingPresets
    {
        public static readonly Dictionary<string, string> RevitExportPresets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "RevitArch", "{name}" },
            { "RevitStruct", "{name}" },
            { "RevitMEP", "{name}" },
            { "SequentialNumber", "{index} - {name}" },
            { "PrefixNumber", "Sheet {index}" },
            { "Architectural", "A{index:D3} - {name}" },
            { "Structural", "S{index:D3} - {name}" },
            { "Mechanical", "M{index:D3} - {name}" },
            { "Electrical", "E{index:D3} - {name}" },
            { "Plumbing", "P{index:D3} - {name}" },
            { "FireProtection", "FP{index:D3} - {name}" },
            { "Civil", "C{index:D3} - {name}" }
        };

        public static string GetPattern(string presetName)
        {
            if (string.IsNullOrEmpty(presetName))
                return null;

            return RevitExportPresets.TryGetValue(presetName, out var pattern) ? pattern : null;
        }
    }
}
