using System.IO;
using System.Runtime.InteropServices;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class HardeningTests
{
    private static readonly PowerPointWarmStartOptions FastWarmStartOptions = new(
        TimeSpan.FromSeconds(1),
        TimeSpan.FromMilliseconds(1),
        TimeSpan.Zero,
        TimeSpan.Zero,
        2,
        TimeSpan.FromSeconds(1));

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
    public void DebugEventsAreWrittenOnlyWhenExplicitlyEnabled()
    {
        var testRoot = CreateTestDirectory();
        try
        {
            var diagnostics = new FileApplicationDiagnostics(testRoot);
            diagnostics.RecordDebugEvent("DisabledTrace", "stage=disabled");
            diagnostics.SetDebugLoggingEnabled(true);
            diagnostics.RecordDebugEvent("EnabledTrace", "stage=activation;elapsedMs=123");

            var log = File.ReadAllText(Path.Combine(testRoot, "PptCompare.log"));
            Assert.IsFalse(log.Contains("DisabledTrace", StringComparison.Ordinal));
            StringAssert.Contains(log, "DebugLoggingEnabled");
            StringAssert.Contains(log, "DEBUG|EnabledTrace|stage=activation;elapsedMs=123");
            StringAssert.Contains(diagnostics.CreateSupportSummary(), "Debug logging: Enabled");
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
    [DataRow(0)]
    [DataRow(1)]
    public async Task ExistingPowerPointSessionRendersWithoutClosingUserWorkOrQuitting(
        int userPresentationCount)
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        await File.WriteAllBytesAsync(presentationPath, [1, 2, 3, 4], TestContext.CancellationToken);
        var detector = new StubPowerPointProcessDetector(isRunning: true);
        var application = new FakePowerPointApplication(userPresentationCount);
        var applicationFactory = new StubPowerPointApplicationFactory(application);
        var launcher = new StubPowerPointProcessLauncher();
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                detector,
                applicationFactory,
                launcher,
                FastWarmStartOptions);

            var result = await renderer.RenderAsync(presentationPath, 1, TestContext.CancellationToken);

            Assert.AreEqual(1, detector.CallCount);
            Assert.AreEqual(1, applicationFactory.CreateCount);
            Assert.AreEqual(0, launcher.StartCount);
            Assert.AreEqual(1, result.SlideImages.Count);
            Assert.IsTrue(application.Presentations.PreviewPresentationClosed);
            Assert.AreEqual(1, application.Presentations.PreviewPresentation.ExportCount);
            Assert.AreEqual(userPresentationCount, application.Presentations.Count);
            Assert.IsFalse(application.QuitCalled);
            Assert.AreEqual(2, application.AutomationSecurity);
            Assert.AreEqual(-1, application.Presentations.ReadOnlyArgument);
            Assert.AreEqual(0, application.Presentations.WithWindowArgument);
            Assert.AreNotEqual(
                Path.GetFullPath(presentationPath),
                application.Presentations.OpenedPath);
            Assert.IsFalse(File.Exists(application.Presentations.OpenedPath));
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task RendererOwnedPowerPointSessionQuitsAfterItsPreviewCloses()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        await File.WriteAllBytesAsync(presentationPath, [1, 2, 3, 4], TestContext.CancellationToken);
        var detector = new StubPowerPointProcessDetector(isRunning: false);
        var application = new FakePowerPointApplication(
            userPresentationCount: 0,
            onQuit: () => detector.SetRunning(false));
        var launcher = new StubPowerPointProcessLauncher(() => detector.SetRunning(true));
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                detector,
                new StubPowerPointApplicationFactory(application),
                launcher,
                FastWarmStartOptions);

            var result = await renderer.RenderAsync(presentationPath, 1, TestContext.CancellationToken);

            Assert.AreEqual(1, result.SlideImages.Count);
            Assert.IsTrue(application.Presentations.PreviewPresentationClosed);
            Assert.AreEqual(0, application.Presentations.Count);
            Assert.IsTrue(application.QuitCalled);
            Assert.AreEqual(1, launcher.StartCount);
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task ColdStartFailureDebugTraceIdentifiesActivationStageWithoutSourceDetails()
    {
        const int serverExecutionFailed = unchecked((int)0x80080005);
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "confidential-board-plan.pptx");
        await File.WriteAllBytesAsync(presentationPath, [1, 2, 3, 4], TestContext.CancellationToken);
        var diagnostics = new RecordingDetailedDiagnostics();
        diagnostics.SetDebugLoggingEnabled(true);
        var activationException = Marshal.GetExceptionForHR(serverExecutionFailed)
            ?? throw new InvalidOperationException("The test HRESULT could not be represented.");
        var detector = new StubPowerPointProcessDetector(isRunning: false);
        var launcher = new StubPowerPointProcessLauncher(() => detector.SetRunning(true));
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                diagnostics,
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                detector,
                new FailingPowerPointApplicationFactory(activationException),
                launcher,
                FastWarmStartOptions);

            var result = await renderer.RenderAsync(presentationPath, 1, TestContext.CancellationToken);

            Assert.AreEqual(0, result.SlideImages.Count);
            StringAssert.Contains(result.Status, "PowerPoint could not start");
            StringAssert.Contains(diagnostics.DebugText, "failedStage=application-activation");
            StringAssert.Contains(diagnostics.DebugText, "hresult=0x80080005");
            StringAssert.Contains(diagnostics.DebugText, "stage=warm-start-process-detected");
            Assert.IsFalse(diagnostics.DebugText.Contains(
                "confidential-board-plan",
                StringComparison.OrdinalIgnoreCase));
            Assert.IsFalse(diagnostics.DebugText.Contains(
                testRoot,
                StringComparison.OrdinalIgnoreCase));
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

    private sealed class StubPowerPointProcessDetector : IPowerPointProcessDetector
    {
        private int _callCount;
        private int _isRunning;

        public StubPowerPointProcessDetector(bool isRunning)
        {
            SetRunning(isRunning);
        }

        public int CallCount => Volatile.Read(ref _callCount);

        public void SetRunning(bool isRunning) =>
            Volatile.Write(ref _isRunning, isRunning ? 1 : 0);

        public bool IsPowerPointRunning()
        {
            Interlocked.Increment(ref _callCount);
            return Volatile.Read(ref _isRunning) != 0;
        }
    }

    private sealed class StubPowerPointProcessLauncher(Action? onStart = null)
        : IPowerPointProcessLauncher
    {
        private int _startCount;

        public int StartCount => Volatile.Read(ref _startCount);

        public void StartPowerPoint()
        {
            Interlocked.Increment(ref _startCount);
            onStart?.Invoke();
        }
    }

    private sealed class StubPowerPointApplicationFactory(FakePowerPointApplication application)
        : IPowerPointApplicationFactory
    {
        private int _createCount;

        public int CreateCount => Volatile.Read(ref _createCount);

        public bool IsPowerPointInstalled() => true;

        public object CreateApplication()
        {
            Interlocked.Increment(ref _createCount);
            return application;
        }
    }

    private sealed class FailingPowerPointApplicationFactory(Exception exception)
        : IPowerPointApplicationFactory
    {
        public bool IsPowerPointInstalled() => true;

        public object CreateApplication() => throw exception;
    }

    private sealed class RecordingDetailedDiagnostics
        : IApplicationDiagnostics, IDetailedApplicationDiagnostics
    {
        private readonly Lock _gate = new();
        private readonly List<string> _debugEvents = [];

        public string DisplayLogLocation => string.Empty;

        public bool IsDebugLoggingEnabled { get; private set; }

        public string DebugText
        {
            get
            {
                lock (_gate)
                {
                    return string.Join(Environment.NewLine, _debugEvents);
                }
            }
        }

        public void SetDebugLoggingEnabled(bool enabled) => IsDebugLoggingEnabled = enabled;

        public void RecordDebugEvent(string eventName, string? nonSensitiveDetail = null)
        {
            if (!IsDebugLoggingEnabled)
            {
                return;
            }

            lock (_gate)
            {
                _debugEvents.Add($"{eventName}|{nonSensitiveDetail}");
            }
        }

        public void RecordEvent(string eventName, string? nonSensitiveDetail = null)
        {
        }

        public void RecordException(string eventName, Exception exception)
        {
        }

        public string CreateSupportSummary() => string.Empty;
    }

    public sealed class FakePowerPointApplication(int userPresentationCount, Action? onQuit = null)
    {
        public int AutomationSecurity { get; set; } = 2;

        public FakePowerPointPresentations Presentations { get; } = new(userPresentationCount);

        public bool QuitCalled { get; private set; }

        public void Quit()
        {
            QuitCalled = true;
            onQuit?.Invoke();
        }
    }

    public sealed class FakePowerPointPresentations(int userPresentationCount)
    {
        private bool _previewPresentationOpen;

        public int Count => userPresentationCount + (_previewPresentationOpen ? 1 : 0);

        public string OpenedPath { get; private set; } = string.Empty;

        public int ReadOnlyArgument { get; private set; }

        public int WithWindowArgument { get; private set; }

        public bool PreviewPresentationClosed { get; private set; }

        public FakePowerPointPresentation PreviewPresentation { get; private set; } = null!;

        public FakePowerPointPresentation Open(
            string path,
            int readOnly,
            int untitled,
            int withWindow)
        {
            OpenedPath = Path.GetFullPath(path);
            ReadOnlyArgument = readOnly;
            WithWindowArgument = withWindow;
            _previewPresentationOpen = true;
            PreviewPresentation = new FakePowerPointPresentation(this);
            return PreviewPresentation;
        }

        public void ClosePreview()
        {
            _previewPresentationOpen = false;
            PreviewPresentationClosed = true;
        }
    }

    public sealed class FakePowerPointPresentation(FakePowerPointPresentations owner)
    {
        public FakePowerPointSlides Slides { get; } = new();

        public FakePowerPointPageSetup PageSetup { get; } = new();

        public int ExportCount { get; private set; }

        public void Export(string folder, string format, int width, int height)
        {
            ExportCount++;
            Directory.CreateDirectory(folder);
            File.WriteAllBytes(Path.Combine(folder, "Slide1.png"), [1, 2, 3, 4]);
        }

        public void Close() => owner.ClosePreview();
    }

    public sealed class FakePowerPointSlides
    {
        public int Count { get; } = 1;
    }

    public sealed class FakePowerPointPageSetup
    {
        public double SlideWidth { get; } = 960;

        public double SlideHeight { get; } = 540;
    }

    public TestContext TestContext { get; set; } = null!;
}
