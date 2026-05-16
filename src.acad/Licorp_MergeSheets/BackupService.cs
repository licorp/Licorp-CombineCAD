using System;
using System.IO;

namespace Licorp_MergeSheets
{
    public class BackupService
    {
        public string CreateBackup(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return null;

            try
            {
                if (!File.Exists(filePath))
                {
                    AcadLogger.LogInfo($"BackupService: No existing file to backup: {filePath}");
                    return null;
                }

                string directory = Path.GetDirectoryName(filePath);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupFileName = $"{fileNameWithoutExt}_backup_{timestamp}{extension}";
                string backupPath = Path.Combine(directory, backupFileName);

                int counter = 1;
                while (File.Exists(backupPath))
                {
                    backupFileName = $"{fileNameWithoutExt}_backup_{timestamp}_{counter}{extension}";
                    backupPath = Path.Combine(directory, backupFileName);
                    counter++;
                }

                File.Copy(filePath, backupPath, false);
                AcadLogger.LogInfo($"BackupService: Created backup: {backupPath}");

                CleanupOldBackups(directory, fileNameWithoutExt, extension, 5);

                return backupPath;
            }
            catch (Exception ex)
            {
                AcadLogger.LogWarning($"BackupService: Failed to backup '{filePath}': {ex.Message}");
                return null;
            }
        }

        private void CleanupOldBackups(string directory, string fileNamePrefix, string extension, int maxBackups)
        {
            try
            {
                string searchPattern = $"{fileNamePrefix}_backup_*{extension}";
                var backups = Directory.GetFiles(directory, searchPattern, SearchOption.TopDirectoryOnly);

                if (backups.Length <= maxBackups)
                    return;

                Array.Sort(backups, StringComparer.OrdinalIgnoreCase);

                int toDelete = backups.Length - maxBackups;
                for (int i = 0; i < toDelete; i++)
                {
                    try
                    {
                        File.Delete(backups[i]);
                        AcadLogger.LogInfo($"BackupService: Cleaned up old backup: {backups[i]}");
                    }
                    catch (Exception ex)
                    {
                        AcadLogger.LogWarning($"BackupService: Failed to delete old backup '{backups[i]}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                AcadLogger.LogWarning($"BackupService: Cleanup failed: {ex.Message}");
            }
        }
    }
}
