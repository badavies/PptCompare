using System.IO;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class HardeningTests
{
    [TestMethod]
    public void ExceptionDiagnosticsExcludeMessagePathAndPresentationName()
    {
        var testRoot = CreateTestDirectory();
        try
        {
            var diagnostics = new FileApplicationDiagnostics(testRoot);
            diagnostics.RecordException(
                "PresentationLoadRejected",
                new InvalidOperationException(
                    @"Private slide text in C:\Sensitive\Board Strategy.pptx"));

            var log = File.ReadAllText(Path.Combine(testRoot, "PptCompare.log"));
            StringAssert.Contains(log, "PresentationLoadRejected");
            StringAssert.Contains(log, "InvalidOperationException");
            Assert.IsFalse(log.Contains("Private slide text", StringComparison.Ordinal));
            Assert.IsFalse(log.Contains("Board Strategy", StringComparison.Ordinal));
            Assert.IsFalse(log.Contains(@"C:\Sensitive", StringComparison.Ordinal));
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public void StartupCleanupDeletesOnlyOldGuidNamedRenderFolders()
    {
        var testRoot = CreateTestDirectory();
        var oldGuidFolder = Path.Combine(testRoot, Guid.NewGuid().ToString("N"));
        var recentGuidFolder = Path.Combine(testRoot, Guid.NewGuid().ToString("N"));
        var unrelatedFolder = Path.Combine(testRoot, "do-not-delete");
        Directory.CreateDirectory(oldGuidFolder);
        Directory.CreateDirectory(recentGuidFolder);
        Directory.CreateDirectory(unrelatedFolder);

        var now = new DateTimeOffset(2026, 9, 12, 12, 0, 0, TimeSpan.Zero);
        Directory.SetLastWriteTimeUtc(oldGuidFolder, now.UtcDateTime.AddDays(-2));
        Directory.SetLastWriteTimeUtc(recentGuidFolder, now.UtcDateTime.AddHours(-2));
        Directory.SetLastWriteTimeUtc(unrelatedFolder, now.UtcDateTime.AddDays(-7));

        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                testRoot,
                new FixedTimeProvider(now));

            Assert.IsFalse(Directory.Exists(oldGuidFolder));
            Assert.IsTrue(Directory.Exists(recentGuidFolder));
            Assert.IsTrue(Directory.Exists(unrelatedFolder));
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public void SupportSummaryContainsEnvironmentButNoUserPaths()
    {
        var testRoot = CreateTestDirectory();
        try
        {
            var diagnostics = new FileApplicationDiagnostics(testRoot);

            var summary = diagnostics.CreateSupportSummary();

            StringAssert.Contains(summary, "Product: PptCompare");
            StringAssert.Contains(summary, "Version: 0.1.0");
            StringAssert.Contains(summary, "presentation content are not recorded");
            Assert.IsFalse(summary.Contains(Environment.UserName, StringComparison.OrdinalIgnoreCase));
            Assert.IsFalse(summary.Contains(testRoot, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    private static string CreateTestDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "PptCompareTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static void DeleteTestDirectory(string path)
    {
        if (Directory.Exists(path))
        {
            Directory.Delete(path, true);
        }
    }

    private sealed class FixedTimeProvider(DateTimeOffset utcNow) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => utcNow;
    }

    private sealed class RecordingDiagnostics : IApplicationDiagnostics
    {
        public string DisplayLogLocation => string.Empty;

        public void RecordEvent(string eventName, string? nonSensitiveDetail = null)
        {
        }

        public void RecordException(string eventName, Exception exception)
        {
        }

        public string CreateSupportSummary() => string.Empty;
    }
}
