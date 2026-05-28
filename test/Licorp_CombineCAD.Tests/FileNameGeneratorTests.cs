using System;
using NUnit.Framework;
using Licorp_CombineCAD.Models;
using Licorp_CombineCAD.Services;

namespace Licorp_CombineCAD.Tests
{
    [TestFixture]
    public class FileNameGeneratorTests
    {
        private SheetInfo CreateTestSheet(string sheetNumber = "A101", string sheetName = "Floor Plan", string paperSize = "A3")
        {
            return new SheetInfo
            {
                SheetNumber = sheetNumber,
                SheetName = sheetName,
                PaperSize = paperSize
            };
        }

        [Test]
        public void GenerateFileName_NullSheet_ReturnsEmpty()
        {
            var result = DwgExportService.GenerateFileName(null, "{SheetNumber}");
            Assert.That(result, Is.EqualTo(""));
        }

        [Test]
        public void GenerateFileName_NullTemplate_UsesDefault()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, null);
            Assert.That(result, Is.EqualTo("A101 - Floor Plan"));
        }

        [Test]
        public void GenerateFileName_EmptyTemplate_UsesDefault()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "");
            Assert.That(result, Is.EqualTo("A101 - Floor Plan"));
        }

        [Test]
        public void GenerateFileName_SheetNumberOnly()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "{SheetNumber}");
            Assert.That(result, Is.EqualTo("A101"));
        }

        [Test]
        public void GenerateFileName_SheetNameOnly()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "{SheetName}");
            Assert.That(result, Is.EqualTo("Floor Plan"));
        }

        [Test]
        public void GenerateFileName_PaperSizeOnly()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "{PaperSize}");
            Assert.That(result, Is.EqualTo("A3"));
        }

        [Test]
        public void GenerateFileName_DateVariable()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "{Date}");
            var expected = DateTime.Now.ToString("yyyyMMdd");
            Assert.That(result, Is.EqualTo(expected));
        }

        [Test]
        public void GenerateFileName_TimeVariable()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "{Time}");
            var expected = DateTime.Now.ToString("HHmm");
            Assert.That(result, Is.EqualTo(expected));
        }

        [Test]
        public void GenerateFileName_MultipleVariables()
        {
            var sheet = CreateTestSheet();
            var result = DwgExportService.GenerateFileName(sheet, "{SheetNumber}_{SheetName}_{PaperSize}");
            Assert.That(result, Is.EqualTo("A101_Floor Plan_A3"));
        }

        [Test]
        public void GenerateFileName_InvalidChars_ReplacedWithDash()
        {
            var sheet = CreateTestSheet("A:101", "Floor/Plan");
            var result = DwgExportService.GenerateFileName(sheet, "{SheetNumber} - {SheetName}");
            Assert.That(result, Does.Not.Contain(":"));
            Assert.That(result, Does.Not.Contain("/"));
        }

        [Test]
        public void GenerateFileName_AllVariables_WithNullValues()
        {
            var sheet = new SheetInfo
            {
                SheetNumber = null,
                SheetName = null,
                PaperSize = null
            };
            var result = DwgExportService.GenerateFileName(sheet, "{SheetNumber} - {SheetName} - {PaperSize}");
            Assert.That(result, Is.EqualTo(" -  - "));
        }

        [Test]
        public void GenerateFileName_ComplexTemplate()
        {
            var sheet = CreateTestSheet("A101", "Floor Plan - Level 1", "A1");
            var template = "{ProjectNumber}_{SheetNumber}_{SheetName}_{PaperSize}_{Date}";
            var result = DwgExportService.GenerateFileName(sheet, template);
            
            Assert.That(result, Does.StartWith("_A101_Floor Plan - Level 1_A1_"));
            Assert.That(result, Does.EndWith(DateTime.Now.ToString("yyyyMMdd")));
        }
    }
}
