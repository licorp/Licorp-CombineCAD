using System.Collections.Generic;
using System.Linq;

namespace Licorp_CombineCAD.Models
{
    public enum PreflightSeverity
    {
        Info,
        Warning,
        Error
    }

    public class SheetPreflightIssue
    {
        public PreflightSeverity Severity { get; set; }
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }
        public string Message { get; set; }

        public string DisplayName
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SheetNumber))
                    return string.IsNullOrWhiteSpace(SheetName) ? "General" : SheetName;

                return string.IsNullOrWhiteSpace(SheetName)
                    ? SheetNumber
                    : SheetNumber + " - " + SheetName;
            }
        }
    }

    public class SheetPreflightResult
    {
        public List<SheetPreflightIssue> Issues { get; set; } = new List<SheetPreflightIssue>();

        public bool HasIssues => Issues.Count > 0;
        public bool HasErrors => Issues.Any(i => i.Severity == PreflightSeverity.Error);
        public bool HasWarnings => Issues.Any(i => i.Severity == PreflightSeverity.Warning);
        public int ErrorCount => Issues.Count(i => i.Severity == PreflightSeverity.Error);
        public int WarningCount => Issues.Count(i => i.Severity == PreflightSeverity.Warning);
        public int InfoCount => Issues.Count(i => i.Severity == PreflightSeverity.Info);

        public string Summary
        {
            get
            {
                if (!HasIssues)
                    return "Preflight passed";

                return string.Format(
                    "Preflight: {0} error(s), {1} warning(s), {2} info",
                    ErrorCount,
                    WarningCount,
                    InfoCount);
            }
        }

        public void AddIssue(PreflightSeverity severity, string sheetNumber, string sheetName, string message)
        {
            Issues.Add(new SheetPreflightIssue
            {
                Severity = severity,
                SheetNumber = sheetNumber,
                SheetName = sheetName,
                Message = message
            });
        }
    }
}
