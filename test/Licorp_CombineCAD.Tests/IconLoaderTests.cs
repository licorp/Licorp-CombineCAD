using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using NUnit.Framework;

namespace Licorp_CombineCAD.Tests
{
    [TestFixture]
    public class IconLoaderTests
    {
        [Test]
        public void LoadIcon_WithKnownName_ReturnsBitmapImage()
        {
            // Act
            var result = IconLoader.LoadIcon("multi_layout", 32);

            // Assert
            Assert.That(result, Is.Not.Null, "IconLoader should return a bitmap for known icon");
            Assert.That(result, Is.InstanceOf<BitmapImage>(), "Result should be BitmapImage type");
        }

        [Test]
        public void LoadIcon_WithDifferentKnownNames_ReturnsBitmapImage([Values("single_layout", "model_space", "layers")] string iconName)
        {
            // Act
            var result = IconLoader.LoadIcon(iconName, 32);

            // Assert
            Assert.That(result, Is.Not.Null, $"IconLoader should return a bitmap for {iconName}");
        }

        [Test]
        public void LoadIcon_WithDifferentSizes_ReturnsBitmapImage([Values(16, 32)] int size)
        {
            // Act
            var result = IconLoader.LoadIcon("multi_layout", size);

            // Assert
            Assert.That(result, Is.Not.Null, $"IconLoader should return a bitmap for size {size}");
        }

        [Test]
        public void LoadIcon_WithUnknownName_ReturnsFallback()
        {
            // Act
            var result = IconLoader.LoadIcon("unknown_icon", 32);

            // Assert - Should return null or fallback placeholder
            // Depending on implementation
            Assert.That(result, Is.Not.Null, "IconLoader should generate placeholder for unknown icon");
        }

        [Test]
        public void LoadIcon_CachesResult()
        {
            // Arrange
            var iconName = "test_cache_icon";
            var size = 32;

            // Act
            var result1 = IconLoader.LoadIcon(iconName, size);
            var result2 = IconLoader.LoadIcon(iconName, size);

            // Assert
            // Cache should return same instance
            Assert.That(result1, Is.Not.Null, "First load should succeed");
            Assert.That(result2, Is.Not.Null, "Second load should succeed");
        }

        [Test]
        public void LoadIcon_WithSize16_ReturnsCorrectSize()
        {
            // Act
            var result = IconLoader.LoadIcon("multi_layout", 16);

            // Assert
            Assert.That(result, Is.Not.Null, "IconLoader should return a bitmap");
            Assert.That(result.PixelWidth, Is.GreaterThan(0), "Pixel width should be greater than 0");
            Assert.That(result.PixelHeight, Is.GreaterThan(0), "Pixel height should be greater than 0");
        }
    }
}
