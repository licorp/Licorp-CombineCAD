using System;
using System.IO;
using System.Linq;
using Nuke.Common;
using Nuke.Common.IO;
using Nuke.Common.ProjectModel;
using Nuke.Common.Tooling;
using Nuke.Common.Utilities.Collections;
using Serilog;

// ReSharper disable CheckNamespace
class Build : NukeBuild
{
    public static int Main() => Execute<Build>(x => x.Compile);

    [Parameter("Configuration to build - Default is 'Debug' (local) or 'Release' (server)")]
    readonly Configuration Configuration = IsServerBuild ? Configuration.Release : Configuration.Debug;

    [Parameter("Target Revit version (e.g., 2020, 2021, etc.) - Default is all")]
    readonly string RevitVersion;

    [Parameter("NuGet API Key")]
    readonly string NuGetApiKey;

    [Solution]
    readonly Solution Solution;

    string[] RevitVersions => new[] { "2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027" };

    Target Clean => _ => _
        .Executes(() =>
        {
            Log.Information("Cleaning output directories...");
            var outputDir = RootDirectory / "bin";
            if (outputDir.Exists())
            {
                try
                {
                    outputDir.DeleteDirectory();
                    Log.Information("Cleaned output directory: {OutputDir}", outputDir);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Could not clean output directory: {OutputDir}", outputDir);
                }
            }
        });

    Target Compile => _ => _
        .DependsOn(Clean)
        .Executes(() =>
        {
            var versionsToBuild = string.IsNullOrEmpty(RevitVersion)
                ? RevitVersions
                : new[] { RevitVersion };

            foreach (var version in versionsToBuild)
            {
                Log.Information("Building for Revit {Version}...", version);

                var projectPath = RootDirectory / "src" / $"Licorp_CombineCAD.R{version}" / $"Licorp_CombineCAD.R{version}.csproj";

                if (!projectPath.Exists())
                {
                    Log.Warning("Project not found: {ProjectPath}", projectPath);
                    continue;
                }

                try
                {
                    DotNetBuild(s => s
                        .SetProjectFile(projectPath)
                        .SetConfiguration(Configuration)
                        .SetVerbosity(DotNetVerbosity.Minimal));

                    Log.Information("Successfully built Revit {Version}", version);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Failed to build Revit {Version}", version);
                    throw;
                }
            }
        });

    Target Pack => _ => _
        .DependsOn(Compile)
        .Executes(() =>
        {
            Log.Information("Packaging bundle...");

            var bundleDir = RootDirectory / "artifacts" / "Licorp_CombineCAD.bundle";
            var contentsDir = bundleDir / "Contents";

            // Clean and create bundle directory
            if (bundleDir.Exists()) bundleDir.DeleteDirectory();
            contentsDir.CreateDirectory();

            // Copy DLLs from all versions
            foreach (var version in RevitVersions)
            {
                var sourceDll = RootDirectory / "bin" / $"R{version}" / Configuration / "Licorp_CombineCAD.dll";
                var targetDir = contentsDir / version;

                if (sourceDll.Exists())
                {
                    targetDir.CreateDirectory();
                    sourceDll.CopyToDirectory(targetDir, FileExistsPolicy.Overwrite);
                    Log.Information("Copied Revit {Version} DLL to bundle", version);
                }
                else
                {
                    Log.Warning("DLL not found for Revit {Version}: {SourceDll}", version, sourceDll);
                }
            }

            // Generate PackageContents.xml
            GeneratePackageContentsXml(bundleDir);

            Log.Information("Bundle packaged at: {BundleDir}", bundleDir);
        });

    Target Deploy => _ => _
        .DependsOn(Pack)
        .Executes(() =>
        {
            Log.Information("Deploying to ApplicationPlugins...");

            var sourceBundle = RootDirectory / "artifacts" / "Licorp_CombineCAD.bundle";
            var targetDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk", "ApplicationPlugins");

            var targetBundle = targetDir / "Licorp_CombineCAD.bundle";

            if (targetBundle.Exists())
            {
                targetBundle.DeleteDirectory();
                Log.Information("Removed existing bundle");
            }

            sourceBundle.Copy(targetBundle);
            Log.Information("Deployed to: {TargetDir}", targetBundle);
        });

    Target Publish => _ => _
        .DependsOn(Pack)
        .Executes(() =>
        {
            Log.Information("Publishing...");

            // Create zip package
            var zipPath = RootDirectory / "artifacts" / $"Licorp_CombineCAD_{DateTime.Now:yyyyMMdd}.zip";
            var bundleDir = RootDirectory / "artifacts" / "Licorp_CombineCAD.bundle";

            CompressionTasks.CompressZip(bundleDir, zipPath);
            Log.Information("Published zip: {ZipPath}", zipPath);
        });

    void GeneratePackageContentsXml(AbsolutePath bundleDir)
    {
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<ApplicationPackage>
  <Name>Licorp_CombineCAD</Name>
  <Description>Combine CAD for Revit</Description>
  <AppName>Licorp CombineCAD</AppName>
  <Version>1.0.0</Version>
  <Components>
";

        foreach (var version in RevitVersions)
        {
            xmlContent += $"    <ComponentEntry Version=\"{version}\" ModuleName=\"Contents/{version}/Licorp_CombineCAD.dll\" />\n";
        }

        xmlContent += @"  </Components>
</ApplicationPackage>";

        var xmlPath = bundleDir / "PackageContents.xml";
        xmlPath.WriteAllText(xmlContent);
        Log.Information("Generated PackageContents.xml");
    }
}
