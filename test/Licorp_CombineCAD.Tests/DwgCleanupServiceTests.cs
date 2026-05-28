using System;
using System.IO;
using NUnit.Framework;
using Licorp_CombineCAD.Services;

namespace Licorp_CombineCAD.Tests
{
    [TestFixture]
    public class DwgCleanupServiceTests
    {
        private string _tempDir;

        [SetUp]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "LicorpTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        [TearDown]
        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_tempDir))
                    Directory.Delete(_tempDir, true);
            }
            catch { }
        }

        [Test]
        public void HasXRefFiles_NullPath_ReturnsFalse()
        {
            var result = DwgCleanupService.HasXRefFiles(null);
            Assert.That(result, Is.False);
        }

        [Test]
        public void HasXRefFiles_EmptyPath_ReturnsFalse()
        {
            var result = DwgCleanupService.HasXRefFiles("");
            Assert.That(result, Is.False);
        }

        [Test]
        public void HasXRefFiles_NonExistentFile_ReturnsFalse()
        {
            var result = DwgCleanupService.HasXRefFiles(@"C:\nonexistent\file.dwg");
            Assert.That(result, Is.False);
        }

        [Test]
        public void HasXRefFiles_NoCompanionFiles_ReturnsFalse()
        {
            var mainFile = Path.Combine(_tempDir, "Main.dwg");
            File.WriteAllText(mainFile, "test content");

            var result = DwgCleanupService.HasXRefFiles(mainFile);
            Assert.That(result, Is.False);
        }

        [Test]
        public void HasXRefFiles_WithCompanionFiles_ReturnsTrue()
        {
            var mainFile = Path.Combine(_tempDir, "Main.dwg");
            var companionFile = Path.Combine(_tempDir, "Main-Model.dwg");
            File.WriteAllText(mainFile, "test content");
            File.WriteAllText(companionFile, "companion content");

            var result = DwgCleanupService.HasXRefFiles(mainFile);
            Assert.That(result, Is.True);
        }

        [Test]
        public void CleanupXRefFiles_NullPath_ReturnsZero()
        {
            var result = DwgCleanupService.CleanupXRefFiles(null);
            Assert.That(result, Is.EqualTo(0));
        }

        [Test]
        public void CleanupXRefFiles_NoCompanionFiles_ReturnsZero()
        {
            var mainFile = Path.Combine(_tempDir, "Main.dwg");
            File.WriteAllText(mainFile, "test content");

            var result = DwgCleanupService.CleanupXRefFiles(mainFile);
            Assert.That(result, Is.EqualTo(0));
        }

        [Test]
        public void CleanupXRefFiles_WithCompanionFiles_DeletesThem()
        {
            var mainFile = Path.Combine(_tempDir, "Main.dwg");
            var companion1 = Path.Combine(_tempDir, "Main-Model.dwg");
            var companion2 = Path.Combine(_tempDir, "Main-Views.dwg");
            File.WriteAllText(mainFile, "test content");
            File.WriteAllText(companion1, "companion 1");
            File.WriteAllText(companion2, "companion 2");

            var result = DwgCleanupService.CleanupXRefFiles(mainFile);
            
            Assert.That(result, Is.EqualTo(2));
            Assert.That(File.Exists(mainFile), Is.True, "Main file should still exist");
            Assert.That(File.Exists(companion1), Is.False, "Companion 1 should be deleted");
            Assert.That(File.Exists(companion2), Is.False, "Companion 2 should be deleted");
        }

        [Test]
        public void GetXRefFiles_NullPath_ReturnsEmpty()
        {
            var result = DwgCleanupService.GetXRefFiles(null);
            Assert.That(result, Is.Empty);
        }

        [Test]
        public void GetXRefFiles_NoCompanionFiles_ReturnsEmpty()
        {
            var mainFile = Path.Combine(_tempDir, "Main.dwg");
            File.WriteAllText(mainFile, "test content");

            var result = DwgCleanupService.GetXRefFiles(mainFile);
            Assert.That(result, Is.Empty);
        }

        [Test]
        public void GetXRefFiles_WithCompanionFiles_ReturnsThem()
        {
            var mainFile = Path.Combine(_tempDir, "Main.dwg");
            var companion1 = Path.Combine(_tempDir, "Main-Model.dwg");
            var companion2 = Path.Combine(_tempDir, "Main-Views.dwg");
            File.WriteAllText(mainFile, "test content");
            File.WriteAllText(companion1, "companion 1");
            File.WriteAllText(companion2, "companion 2");

            var result = DwgCleanupService.GetXRefFiles(mainFile);
            
            Assert.That(result, Has.Length.EqualTo(2));
            Assert.That(result, Does.Contain(companion1));
            Assert.That(result, Does.Contain(companion2));
        }
    }
}
