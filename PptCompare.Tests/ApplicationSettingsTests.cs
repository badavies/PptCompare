using PptCompare.Models;

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
        Assert.AreEqual(25 * 1024 * 1024, limits.MaxPreviewImageBytes);
        Assert.IsFalse(settings.EnableDebugLogging);
    }

    [TestMethod]
    public void UnsafeOrInconsistentLimitsAreRejected()
    {
        Assert.IsFalse(ApplicationSettings.TryValidate(
            new ApplicationSettings(MaxFileSizeMb: 1_001),
            out _));
        Assert.IsFalse(ApplicationSettings.TryValidate(
            new ApplicationSettings(MaxPreviewImageSizeMb: 50, MaxTotalPreviewImageSizeMb: 25),
            out _));
        Assert.IsFalse(ApplicationSettings.TryValidate(
            new ApplicationSettings(OutputFontSizePoints: double.NaN),
            out _));
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
        Assert.AreEqual("0.1.0", ApplicationInfo.Version);
    }
}
