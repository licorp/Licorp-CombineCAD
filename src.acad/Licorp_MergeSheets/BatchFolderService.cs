using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Licorp_MergeSheets
{
    public class BatchFolderService
    {
        public List<SourceFile> ScanFolder(string folderPath, string pattern = "*.dwg", bool recursive = false)
        {
            var result = new List<SourceFile>();

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                AcadLogger.LogWarning($"BatchFolderService: Folder not found: {folderPath}");
                return result;
            }

            var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            string[] patterns = string.IsNullOrEmpty(pattern)
                ? new[] { "*.dwg" }
                : pattern.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

            var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var pat in patterns)
            {
                string trimmedPattern = pat.Trim();
                if (string.IsNullOrEmpty(trimmedPattern))
                    continue;

                try
                {
                    var files = Directory.GetFiles(folderPath, trimmedPattern, searchOption);
                    foreach (var file in files)
                    {
                        string fullPath = Path.GetFullPath(file);
                        if (seenPaths.Contains(fullPath))
                            continue;

                        seenPaths.Add(fullPath);
                        result.Add(new SourceFile
                        {
                            Path = fullPath,
                            Layout = null
                        });
                    }
                }
                catch (Exception ex)
                {
                    AcadLogger.LogWarning($"BatchFolderService: Error scanning pattern '{trimmedPattern}': {ex.Message}");
                }
            }

            result = result.OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase).ToList();
            AcadLogger.LogInfo($"BatchFolderService: Found {result.Count} DWG file(s) in '{folderPath}' (pattern='{pattern}', recursive={recursive})");

            return result;
        }

        public List<SourceFile> ScanFolderWithLayouts(string folderPath, string pattern = "*.dwg", bool recursive = false)
        {
            var files = ScanFolder(folderPath, pattern, recursive);

            foreach (var file in files)
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(file.Path);
                    file.Layout = SanitizeLayoutNameFromFileName(fileName);
                }
                catch (Exception ex)
                {
                    AcadLogger.LogWarning($"BatchFolderService: Error extracting layout name from '{file.Path}': {ex.Message}");
                }
            }

            return files;
        }

        private string SanitizeLayoutNameFromFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "Layout";

            var invalidChars = new HashSet<char>("<>/\\\":;?*|=,&".ToCharArray());
            var chars = fileName
                .Trim()
                .Select(c => invalidChars.Contains(c) || char.IsControl(c) ? ' ' : c)
                .ToArray();

            var safeName = new string(chars).Trim();
            while (safeName.Contains("  "))
                safeName = safeName.Replace("  ", " ");

            if (string.IsNullOrWhiteSpace(safeName))
                safeName = "Layout";

            if (safeName.Length > 31)
                safeName = safeName.Substring(0, 31).TrimEnd();

            return safeName;
        }
    }
}
