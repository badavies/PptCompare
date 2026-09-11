using System.IO;
using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class PresentationSourceSecurityTests
{
    [TestMethod]
    public async Task NonPowerPointExtensionIsRejected()
    {
        var path = CreateTestPath("not-a-presentation.txt");
        try
        {
            await File.WriteAllTextAsync(path, "plain text");
            var source = new OpenXmlPresentationSourceService();

            await Assert.ThrowsExactlyAsync<PresentationLoadException>(() =>
                source.LoadAsync(new PresentationReference(path, "not-a-presentation.txt")));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task CorruptOpenXmlPackageIsRejected()
    {
        var path = CreateTestPath("corrupt.pptx");
        try
        {
            await File.WriteAllTextAsync(path, "not an Open XML package");
            var source = new OpenXmlPresentationSourceService();

            await Assert.ThrowsExactlyAsync<PresentationLoadException>(() =>
                source.LoadAsync(new PresentationReference(path, "corrupt.pptx")));
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string CreateTestPath(string fileName)
    {
        var folder = Path.Combine(Path.GetTempPath(), "PptCompareTests");
        Directory.CreateDirectory(folder);
        return Path.Combine(folder, $"{Guid.NewGuid():N}-{fileName}");
    }
}
