using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Licorp_CombineCAD.Services
{
    public class DwgPostProcessOptions
    {
        public bool CombineAdjacentText { get; set; }
        public bool ConvertRasterToOle { get; set; }
        public bool UpdateScaleLabels { get; set; }
        public bool FlattenZ { get; set; }
        public bool PurgeDrawing { get; set; }
        public bool StripTextFormatting { get; set; }
        public bool AuditDrawing { get; set; }
    }

    public class DwgPostProcessService
    {
        private readonly string _accoreconsolePath;

        public DwgPostProcessService(string accoreconsolePath = null)
        {
            _accoreconsolePath = accoreconsolePath ?? AutoCadLocatorService.FindAcCoreConsole();
        }

        public bool IsAvailable => !string.IsNullOrEmpty(_accoreconsolePath);

        public bool HasAnyOption(DwgPostProcessOptions options)
        {
            return options.CombineAdjacentText || options.ConvertRasterToOle ||
                   options.UpdateScaleLabels || options.FlattenZ ||
                   options.PurgeDrawing || options.StripTextFormatting ||
                   options.AuditDrawing;
        }

        public bool PostProcess(string dwgPath, DwgPostProcessOptions options, int timeoutMs = 180000)
        {
            if (!IsAvailable || !File.Exists(dwgPath))
                return false;

            if (!HasAnyOption(options))
                return true;

            try
            {
                var scriptPath = Path.Combine(Path.GetTempPath(), $"PostProc_{Guid.NewGuid()}.scr");
                var sb = new StringBuilder();

                sb.AppendLine("_SECURELOAD 0");

                if (options.AuditDrawing)
                {
                    sb.AppendLine("_AUDIT");
                    sb.AppendLine("_YES");
                    sb.AppendLine("-PURGE");
                    sb.AppendLine("A");
                    sb.AppendLine("*");
                    sb.AppendLine("N");
                }

                if (options.FlattenZ)
                {
                    sb.AppendLine("(vl-load-com)");
                    sb.AppendLine("(defun c:FlattenZ (/ ss i ent ed pt10 pt11)");
                    sb.AppendLine("  (setq ss (ssget \"X\" '((0 . \"LINE,ARC,CIRCLE,ELLIPSE,LWPOLYLINE,POINT,TEXT,MTEXT,DIMENSION,LEADER,MULTILEADER,HATCH,INSERT\"))))");
                    sb.AppendLine("  (if ss");
                    sb.AppendLine("    (progn");
                    sb.AppendLine("      (setq i 0 cnt 0)");
                    sb.AppendLine("      (while (< i (sslength ss))");
                    sb.AppendLine("        (setq ent (entget (ssname ss i)))");
                    sb.AppendLine("        (if (assoc 10 ent)");
                    sb.AppendLine("          (progn");
                    sb.AppendLine("            (setq pt10 (assoc 10 ent))");
                    sb.AppendLine("            (if (not (equal (caddr (cdr pt10)) 0.0 1e-6))");
                    sb.AppendLine("              (progn");
                    sb.AppendLine("                (setq ent (subst (cons 10 (list (cadr (cdr pt10)) (caddr (cdr pt10)) 0.0)) pt10 ent))");
                    sb.AppendLine("                (setq cnt (1+ cnt))");
                    sb.AppendLine("              )");
                    sb.AppendLine("            )");
                    sb.AppendLine("          )");
                    sb.AppendLine("        )");
                    sb.AppendLine("        (if (assoc 11 ent)");
                    sb.AppendLine("          (progn");
                    sb.AppendLine("            (setq pt11 (assoc 11 ent))");
                    sb.AppendLine("            (if (not (equal (caddr (cdr pt11)) 0.0 1e-6))");
                    sb.AppendLine("              (progn");
                    sb.AppendLine("                (setq ent (subst (cons 11 (list (cadr (cdr pt11)) (caddr (cdr pt11)) 0.0)) pt11 ent))");
                    sb.AppendLine("                (setq cnt (1+ cnt))");
                    sb.AppendLine("              )");
                    sb.AppendLine("            )");
                    sb.AppendLine("          )");
                    sb.AppendLine("        )");
                    sb.AppendLine("        (entmod ent)");
                    sb.AppendLine("        (setq i (1+ i))");
                    sb.AppendLine("      )");
                    sb.AppendLine("      (princ (strcat \"\\nFlattened \" (itoa cnt) \" points to Z=0\"))");
                    sb.AppendLine("    )");
                    sb.AppendLine("  )");
                    sb.AppendLine(")");
                    sb.AppendLine("(c:FlattenZ)");
                }

                if (options.StripTextFormatting)
                {
                    sb.AppendLine("(vl-load-com)");
                    sb.AppendLine("(defun c:StripText (/ ss i ent txt newTxt cnt)");
                    sb.AppendLine("  (setq ss (ssget \"X\" '((0 . \"MTEXT\"))) cnt 0)");
                    sb.AppendLine("  (if ss");
                    sb.AppendLine("    (progn");
                    sb.AppendLine("      (setq i 0)");
                    sb.AppendLine("      (while (< i (sslength ss))");
                    sb.AppendLine("        (setq ent (entget (ssname ss i)))");
                    sb.AppendLine("        (setq txt (cdr (assoc 1 ent)))");
                    sb.AppendLine("        (if (wcmatch txt \"*\\\\*\")");
                    sb.AppendLine("          (progn");
                    sb.AppendLine("            (setq newTxt (cdr (assoc 3 ent)))");
                    sb.AppendLine("            (if (and newTxt (/= newTxt txt))");
                    sb.AppendLine("              (progn");
                    sb.AppendLine("                (entmod (subst (cons 1 newTxt) (assoc 1 ent) ent))");
                    sb.AppendLine("                (setq cnt (1+ cnt))");
                    sb.AppendLine("              )");
                    sb.AppendLine("            )");
                    sb.AppendLine("          )");
                    sb.AppendLine("        )");
                    sb.AppendLine("        (setq i (1+ i))");
                    sb.AppendLine("      )");
                    sb.AppendLine("      (princ (strcat \"\\nStripped formatting from \" (itoa cnt) \" MText objects\"))");
                    sb.AppendLine("    )");
                    sb.AppendLine("  )");
                    sb.AppendLine(")");
                    sb.AppendLine("(c:StripText)");
                }

                if (options.CombineAdjacentText)
                {
                    sb.AppendLine("(vl-load-com)");
                    sb.AppendLine("(defun c:CombineText ()");
                    sb.AppendLine(" (setq ss (ssget \"X\" '((0 . \"TEXT,MTEXT\"))))");
                    sb.AppendLine(" (if ss");
                    sb.AppendLine(" (progn");
                    sb.AppendLine(" (setq i 0 cnt 0)");
                    sb.AppendLine(" (while (< i (sslength ss))");
                    sb.AppendLine(" (setq ent (entget (ssname ss i)))");
                    sb.AppendLine(" (setq pt1 (cdr (assoc 10 ent)))");
                    sb.AppendLine(" (setq txt1 (cdr (assoc 1 ent)))");
                    sb.AppendLine(" (setq j (1+ i) found nil)");
                    sb.AppendLine(" (while (and (not found) (< j (sslength ss)))");
                    sb.AppendLine(" (setq ent2 (entget (ssname ss j)))");
                    sb.AppendLine(" (setq pt2 (cdr (assoc 10 ent2)))");
                    sb.AppendLine(" (setq txt2 (cdr (assoc 1 ent2)))");
                    sb.AppendLine(" (setq dist (distance pt1 pt2))");
                    sb.AppendLine(" (if (< dist 5.0)");
                    sb.AppendLine(" (progn");
                    sb.AppendLine(" (setq newTxt (strcat txt1 \" \" txt2))");
                    sb.AppendLine(" (entmod (subst (cons 1 newTxt) (assoc 1 ent) ent))");
                    sb.AppendLine(" (entdel (ssname ss j))");
                    sb.AppendLine(" (setq cnt (1+ cnt) found T) ) )");
                    sb.AppendLine(" (setq j (1+ j)) )");
                    sb.AppendLine(" (setq i (1+ i)) )");
                    sb.AppendLine(" (princ (strcat \"\\nCombined \" (itoa cnt) \" text objects\")) ) )");
                    sb.AppendLine(")");
                    sb.AppendLine("(c:CombineText)");
                }

                if (options.ConvertRasterToOle)
                {
                    sb.AppendLine("(vl-load-com)");
                    sb.AppendLine("(defun c:RasterToOLE ()");
                    sb.AppendLine(" (setq ss (ssget \"X\" '((0 . \"IMAGE\"))))");
                    sb.AppendLine(" (if ss");
                    sb.AppendLine(" (progn");
                    sb.AppendLine(" (setq i 0 cnt 0)");
                    sb.AppendLine(" (while (< i (sslength ss))");
                    sb.AppendLine(" (setq img (ssname ss i))");
                    sb.AppendLine(" (setq imgObj (vlax-ename->vla-object img))");
                    sb.AppendLine(" (setq imgFile (vla-get-ImageFile imgObj))");
                    sb.AppendLine(" (princ (strcat \"\\nImage: \" imgFile))");
                    sb.AppendLine(" (setq cnt (1+ cnt))");
                    sb.AppendLine(" (setq i (1+ i)) )");
                    sb.AppendLine(" (princ (strcat \"\\nFound \" (itoa cnt) \" raster images\")) ) )");
                    sb.AppendLine(")");
                    sb.AppendLine("(c:RasterToOLE)");
                }

                if (options.UpdateScaleLabels)
                {
                    sb.AppendLine("(vl-load-com)");
                    sb.AppendLine("(defun c:FixScaleLabels ()");
                    sb.AppendLine(" (setq ss (ssget \"X\" '((0 . \"TEXT\") (1 . \"*1:*\"))))");
                    sb.AppendLine(" (if ss");
                    sb.AppendLine(" (progn");
                    sb.AppendLine(" (setq i 0 cnt 0)");
                    sb.AppendLine(" (while (< i (sslength ss))");
                    sb.AppendLine(" (setq ent (entget (ssname ss i)))");
                    sb.AppendLine(" (setq txt (cdr (assoc 1 ent)))");
                    sb.AppendLine(" (if (wcmatch txt \"1:*\") (setq cnt (1+ cnt)))");
                    sb.AppendLine(" (setq i (1+ i)) )");
                    sb.AppendLine(" (princ (strcat \"\\nFound \" (itoa cnt) \" scale labels\")) ) )");
                    sb.AppendLine(")");
                    sb.AppendLine("(c:FixScaleLabels)");
                }

                if (options.PurgeDrawing)
                {
                    sb.AppendLine("-PURGE");
                    sb.AppendLine("A");
                    sb.AppendLine("*");
                    sb.AppendLine("N");
                    sb.AppendLine("-PURGE");
                    sb.AppendLine("A");
                    sb.AppendLine("*");
                    sb.AppendLine("N");
                    sb.AppendLine("-PURGE");
                    sb.AppendLine("R");
                    sb.AppendLine("N");
                    sb.AppendLine("._-SCALELISTEDIT");
                    sb.AppendLine("R");
                    sb.AppendLine("Y");
                    sb.AppendLine("E");
                }

                sb.AppendLine("QSAVE");

                File.WriteAllText(scriptPath, sb.ToString());

                var success = RunAcCoreConsole(dwgPath, scriptPath, timeoutMs);

                try { File.Delete(scriptPath); } catch { }

                return success;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PostProcess] Error: {ex.Message}");
                return false;
            }
        }

        private bool RunAcCoreConsole(string inputPath, string scriptPath, int timeoutMs)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = _accoreconsolePath,
                    Arguments = $"/i \"{inputPath}\" /s \"{scriptPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(startInfo))
                {
                    if (process == null) return false;

                    process.WaitForExit(timeoutMs);

                    if (!process.HasExited)
                    {
                        try { process.Kill(); } catch { }
                        return false;
                    }

                    Debug.WriteLine($"[PostProcess] AcCoreConsole exit: {process.ExitCode}");
                    return process.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PostProcess] Console error: {ex.Message}");
                return false;
            }
    }
    }
}
