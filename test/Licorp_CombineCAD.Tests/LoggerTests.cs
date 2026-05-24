using System.IO;
using System;
using NUnit.Framework;

namespace Licorp_CombineCAD.Tests
{
    [TestFixture]
    public class LoggerTests
    {
        [SetUp]
        public void Setup()
        {
            // Each test gets fresh logger state
        }

        [Test]
        public void Initialize_CreatesLogDirectory()
        {
            // Act
            Logger.Initialize();

            // Assert
            var expectedDir = GetLogDirectory();
            Assert.That(Directory.Exists(expectedDir), Is.True, "Log directory should be created");
        }

        [Test]
        public void LogInfo_DoesNotThrow()
        {
            // Arrange
            Logger.Initialize();

            // Act & Assert
            Assert.DoesNotThrow(() => 
                Logger.LogInfo("Test information message"),
                "LogInfo should not throw");
        }

        [Test]
        public void LogDebug_DoesNotThrow()
        {
            // Arrange
            Logger.Initialize();

            // Act & Assert
            Assert.DoesNotThrow(() => 
                Logger.LogDebug("Test debug message"),
                "LogDebug should not throw");
        }

        [Test]
        public void LogWarning_DoesNotThrow()
        {
            // Arrange
            Logger.Initialize();

            // Act & Assert
            Assert.DoesNotThrow(() => 
                Logger.LogWarning("Test warning message"),
                "LogWarning should not throw");
        }

        [Test]
        public void LogError_DoesNotThrow()
        {
            // Arrange
            Logger.Initialize();

            // Act & Assert
            Assert.DoesNotThrow(() => 
                Logger.LogError("Test error message"),
                "LogError should not throw");
        }

        [Test]
        public void LogSection_DoesNotThrow()
        {
            // Arrange
            Logger.Initialize();

            // Act & Assert
            Assert.DoesNotThrow(() => 
                Logger.LogSection("Test Section"),
                "LogSection should not throw");
        }

        [Test]
        public void LogException_DoesNotThrow()
        {
            // Arrange
            Logger.Initialize();
            var exception = new InvalidOperationException("Test exception");

            // Act & Assert
            Assert.DoesNotThrow(() => 
                Logger.LogException(exception, "Test context"),
                "LogException should not throw");
        }

        [Test]
        public void GetLogFilePath_ReturnsPathOrNotInitialized()
        {
            // Arrange
            Logger.Initialize();

            // Act
            var path = Logger.GetLogFilePath();

            // Assert
            Assert.That(path, Is.Not.Null, "GetLogFilePath should return a value");
            Assert.That(path, Is.Not.Empty, "GetLogFilePath should not return empty string");
        }

        [Test]
        public void GetBufferedLog_ReturnsString()
        {
            // Arrange
            Logger.Initialize();

            // Act
            var log = Logger.GetBufferedLog();

            // Assert
            Assert.That(log, Is.Not.Null, "GetBufferedLog should return a value");
        }

        private static string GetLogDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Licorp_CombineCAD", "Logs");
        }
    }
}
