using System.IO;
using System.Diagnostics.CodeAnalysis;
using PptCompare.Models;
using PptCompare.Services;
using PptCompare.ViewModels;

namespace PptCompare.Tests;

[TestClass]
public sealed class ApplicationSettingsTests
{
    [TestMethod]
    public void DefaultsAreValidAndMapToReadLimits()
    {
        var settings = new ApplicationSettings();

        Assert.IsTrue(ApplicationSettings.TryValidate(settings, out var error), error);
        var limits = settings.ToReadLimits();
        Assert.AreEqual(250L * 1024 * 1024, limits.MaxFileBytes);
        Assert.AreEqual(2_000, limits.MaxSlides);
        Assert.AreEqual(100_000, limits.MaxRelationshipReferences);
        Assert.AreEqual(10_000, limits.MaxUniqueRelatedParts);
        Assert.AreEqual(500L * 1024 * 1024, limits.MaxTotalRelatedPartBytes);
        Assert.AreEqual(25 * 1024 * 1024, limits.MaxPreviewImageBytes);
        Assert.IsFalse(settings.EnableDebugLogging);
    }

    [TestMethod]
    public void EveryNumericSettingAcceptsItsInclusiveBoundaries()
    {
        var validBoundaryCases = new (string Name, ApplicationSettings Settings)[]
        {
            ("minimum output font size", new ApplicationSettings(OutputFontSizePoints: 8)),
            ("maximum output font size", new ApplicationSettings(OutputFontSizePoints: 24)),
            ("minimum file size", new ApplicationSettings(MaxFileSizeMb: 1)),
            ("maximum file size", new ApplicationSettings(MaxFileSizeMb: 1_000)),
            ("minimum slides", new ApplicationSettings(MaxSlides: 1)),
            ("maximum slides", new ApplicationSettings(MaxSlides: 5_000)),
            ("minimum elements", new ApplicationSettings(MaxElementsPerSlide: 1)),
            ("maximum elements", new ApplicationSettings(MaxElementsPerSlide: 50_000)),
            ("minimum table cells", new ApplicationSettings(MaxTableCellsPerSlide: 1)),
            ("maximum table cells", new ApplicationSettings(MaxTableCellsPerSlide: 50_000)),
            ("minimum paragraphs", new ApplicationSettings(MaxParagraphsPerSlide: 1)),
            ("maximum paragraphs", new ApplicationSettings(MaxParagraphsPerSlide: 50_000)),
            ("minimum slide characters", new ApplicationSettings(MaxCharactersPerSlide: 1_000)),
            ("maximum slide characters", new ApplicationSettings(MaxCharactersPerSlide: 2_000_000)),
            (
                "minimum total characters",
                new ApplicationSettings(MaxCharactersPerSlide: 1_000, MaxTotalCharacters: 1_000)),
            ("maximum total characters", new ApplicationSettings(MaxTotalCharacters: 100_000_000)),
            ("minimum embedded item", new ApplicationSettings(MaxEmbeddedItemSizeMb: 1)),
            ("maximum embedded item", new ApplicationSettings(MaxEmbeddedItemSizeMb: 500)),
            ("minimum preview image", new ApplicationSettings(MaxPreviewImageSizeMb: 1)),
            ("maximum preview image", new ApplicationSettings(MaxPreviewImageSizeMb: 100)),
            (
                "minimum total preview",
                new ApplicationSettings(MaxPreviewImageSizeMb: 1, MaxTotalPreviewImageSizeMb: 1)),
            ("maximum total preview", new ApplicationSettings(MaxTotalPreviewImageSizeMb: 500))
        };

        foreach (var (name, settings) in validBoundaryCases)
        {
            Assert.IsTrue(
                ApplicationSettings.TryValidate(settings, out var error),
                $"{name} should be valid: {error}");
        }
    }

    [TestMethod]
    public void EveryNumericSettingRejectsValuesOutsideItsBoundaries()
    {
        var invalidBoundaryCases = new (string Name, ApplicationSettings Settings)[]
        {
            ("output font below minimum", new ApplicationSettings(OutputFontSizePoints: 7.99)),
            ("output font above maximum", new ApplicationSettings(OutputFontSizePoints: 24.01)),
            ("non-finite output font", new ApplicationSettings(OutputFontSizePoints: double.NaN)),
            ("infinite output font", new ApplicationSettings(OutputFontSizePoints: double.PositiveInfinity)),
            ("file size below minimum", new ApplicationSettings(MaxFileSizeMb: 0)),
            ("file size above maximum", new ApplicationSettings(MaxFileSizeMb: 1_001)),
            ("slides below minimum", new ApplicationSettings(MaxSlides: 0)),
            ("slides above maximum", new ApplicationSettings(MaxSlides: 5_001)),
            ("elements below minimum", new ApplicationSettings(MaxElementsPerSlide: 0)),
            ("elements above maximum", new ApplicationSettings(MaxElementsPerSlide: 50_001)),
            ("table cells below minimum", new ApplicationSettings(MaxTableCellsPerSlide: 0)),
            ("table cells above maximum", new ApplicationSettings(MaxTableCellsPerSlide: 50_001)),
            ("paragraphs below minimum", new ApplicationSettings(MaxParagraphsPerSlide: 0)),
            ("paragraphs above maximum", new ApplicationSettings(MaxParagraphsPerSlide: 50_001)),
            ("slide characters below minimum", new ApplicationSettings(MaxCharactersPerSlide: 999)),
            ("slide characters above maximum", new ApplicationSettings(MaxCharactersPerSlide: 2_000_001)),
            (
                "total characters below slide limit",
                new ApplicationSettings(MaxCharactersPerSlide: 10_000, MaxTotalCharacters: 9_999)),
            ("total characters above maximum", new ApplicationSettings(MaxTotalCharacters: 100_000_001)),
            ("embedded item below minimum", new ApplicationSettings(MaxEmbeddedItemSizeMb: 0)),
            ("embedded item above maximum", new ApplicationSettings(MaxEmbeddedItemSizeMb: 501)),
            ("preview image below minimum", new ApplicationSettings(MaxPreviewImageSizeMb: 0)),
            ("preview image above maximum", new ApplicationSettings(MaxPreviewImageSizeMb: 101)),
            (
                "total preview below image limit",
                new ApplicationSettings(MaxPreviewImageSizeMb: 50, MaxTotalPreviewImageSizeMb: 49)),
            ("total preview above maximum", new ApplicationSettings(MaxTotalPreviewImageSizeMb: 501))
        };

        foreach (var (name, settings) in invalidBoundaryCases)
        {
            Assert.IsFalse(
                ApplicationSettings.TryValidate(settings, out _),
                $"{name} should be rejected.");
        }

        Assert.IsFalse(ApplicationSettings.TryValidate(null, out _));
    }

    [TestMethod]
    public void SettingsEditorRejectsIncorrectInputTypes()
    {
        var invalidInputs = new (string Name, Action<SettingsWindowViewModel> SetInvalidValue)[]
        {
            ("output font", viewModel => viewModel.OutputFontSizePoints = "not-a-number"),
            ("file size", viewModel => viewModel.MaxFileSizeMb = "1.5"),
            ("slides", viewModel => viewModel.MaxSlides = "many"),
            ("elements", viewModel => viewModel.MaxElementsPerSlide = "many"),
            ("table cells", viewModel => viewModel.MaxTableCellsPerSlide = "many"),
            ("paragraphs", viewModel => viewModel.MaxParagraphsPerSlide = "many"),
            ("slide characters", viewModel => viewModel.MaxCharactersPerSlide = "many"),
            ("total characters", viewModel => viewModel.MaxTotalCharacters = "many"),
            ("embedded item", viewModel => viewModel.MaxEmbeddedItemSizeMb = "many"),
            ("preview image", viewModel => viewModel.MaxPreviewImageSizeMb = "many"),
            ("total preview", viewModel => viewModel.MaxTotalPreviewImageSizeMb = "many")
        };

        foreach (var (name, setInvalidValue) in invalidInputs)
        {
            var viewModel = new SettingsWindowViewModel(new ApplicationSettings());
            setInvalidValue(viewModel);

            Assert.IsFalse(
                viewModel.TryCreateSettings(out var settings, out var error),
                $"{name} should reject the incorrect input type.");
            Assert.IsNull(settings);
            Assert.IsFalse(string.IsNullOrWhiteSpace(error));
        }
    }

    [TestMethod]
    public async Task InvalidSettingsJsonFallsBackToDefaults()
    {
        var invalidDocuments = new (string Name, string Json)[]
        {
            ("malformed", "{ not-json"),
            ("unknown property", "{\"UnexpectedSetting\":true}"),
            ("wrong numeric type", "{\"MaxSlides\":\"many\"}"),
            ("wrong boolean type", "{\"UsePowerPointRendering\":1}"),
            ("out of range", "{\"MaxSlides\":0}")
        };

        foreach (var (name, json) in invalidDocuments)
        {
            var (folder, path) = CreateIsolatedSettingsPath();
            try
            {
                Directory.CreateDirectory(folder);
                await File.WriteAllTextAsync(path, json, TestContext.CancellationToken);

                var loaded = new JsonApplicationSettingsService(path).Load();

                Assert.AreEqual(new ApplicationSettings(), loaded, $"{name} JSON should be ignored.");
            }
            finally
            {
                Directory.Delete(folder, true);
            }
        }
    }

    [TestMethod]
    public async Task EmptyAndOversizedSettingsFilesFallBackToDefaults()
    {
        var documents = new (string Name, string Contents)[]
        {
            ("empty", string.Empty),
            ("oversized", new string(' ', (64 * 1024) + 1))
        };

        foreach (var (name, contents) in documents)
        {
            var (folder, path) = CreateIsolatedSettingsPath();
            try
            {
                Directory.CreateDirectory(folder);
                await File.WriteAllTextAsync(path, contents, TestContext.CancellationToken);

                var loaded = new JsonApplicationSettingsService(path).Load();

                Assert.AreEqual(new ApplicationSettings(), loaded, $"{name} settings should be ignored.");
            }
            finally
            {
                Directory.Delete(folder, true);
            }
        }
    }

    [TestMethod]
    public void ValidSettingsRoundTripAndInvalidSettingsCannotBeSaved()
    {
        var (folder, path) = CreateIsolatedSettingsPath();
        try
        {
            var expected = new ApplicationSettings(
                OutputFontSizePoints: 11,
                MaxSlides: 750,
                UsePowerPointRendering: false,
                EnableDebugLogging: true);
            var service = new JsonApplicationSettingsService(path);

            service.Save(expected);

            Assert.AreEqual(expected, service.Load());
            Assert.ThrowsExactly<ArgumentException>(() =>
                service.Save(new ApplicationSettings(MaxSlides: 0)));
        }
        finally
        {
            Directory.Delete(folder, true);
        }
    }

    [TestMethod]
    public void ProductMetadataContainsRequiredOwnershipDetails()
    {
        var infoType = typeof(ApplicationInfo);
        var productName = infoType.GetField(nameof(ApplicationInfo.ProductName))?.GetRawConstantValue();
        var designer = infoType.GetField(nameof(ApplicationInfo.Designer))?.GetRawConstantValue();
        var copyright = infoType.GetField(nameof(ApplicationInfo.Copyright))?.GetRawConstantValue();

        Assert.AreEqual("PptCompare", productName);
        Assert.AreEqual("Ben Davies", designer);
        Assert.AreEqual("© 2026 Ben Davies", copyright);
        Assert.AreEqual("0.3.0", ApplicationInfo.Version);
    }

    private static (string Folder, string Path) CreateIsolatedSettingsPath()
    {
        var folder = Path.Combine(Path.GetTempPath(), "PptCompareTests", Guid.NewGuid().ToString("N"));
        return (folder, Path.Combine(folder, "settings.json"));
    }

    [SuppressMessage(
        "ReSharper",
        "MemberCanBePrivate.Global",
        Justification = "MSTest requires a public TestContext property for runtime injection.")]
    [SuppressMessage(
        "ReSharper",
        "AutoPropertyCanBeMadeGetOnly.Global",
        Justification = "MSTest sets TestContext through this property at runtime.")]
    public TestContext TestContext { get; set; } = null!;
}
