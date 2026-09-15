using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Win32;
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
    public void StartupCleanupDeletesOnlyOldGuidNamedManagedFolders()
    {
        var testRoot = CreateTestDirectory();
        var renderRoot = Path.Combine(testRoot, "renders");
        var stagingRoot = Path.Combine(testRoot, "sources");
        var oldGuidFolder = Path.Combine(renderRoot, Guid.NewGuid().ToString("N"));
        var recentGuidFolder = Path.Combine(renderRoot, Guid.NewGuid().ToString("N"));
        var unrelatedFolder = Path.Combine(renderRoot, "do-not-delete");
        var oldStagingFolder = Path.Combine(stagingRoot, Guid.NewGuid().ToString("N"));
        var recentStagingFolder = Path.Combine(stagingRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(oldGuidFolder);
        Directory.CreateDirectory(recentGuidFolder);
        Directory.CreateDirectory(unrelatedFolder);
        Directory.CreateDirectory(oldStagingFolder);
        Directory.CreateDirectory(recentStagingFolder);

        var now = new DateTimeOffset(2026, 9, 12, 12, 0, 0, TimeSpan.Zero);
        Directory.SetLastWriteTimeUtc(oldGuidFolder, now.UtcDateTime.AddDays(-2));
        Directory.SetLastWriteTimeUtc(recentGuidFolder, now.UtcDateTime.AddHours(-2));
        Directory.SetLastWriteTimeUtc(unrelatedFolder, now.UtcDateTime.AddDays(-7));
        Directory.SetLastWriteTimeUtc(oldStagingFolder, now.UtcDateTime.AddDays(-2));
        Directory.SetLastWriteTimeUtc(recentStagingFolder, now.UtcDateTime.AddHours(-2));

        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                renderRoot,
                new FixedTimeProvider(now),
                stagingRoot: stagingRoot);

            Assert.IsFalse(Directory.Exists(oldGuidFolder));
            Assert.IsTrue(Directory.Exists(recentGuidFolder));
            Assert.IsTrue(Directory.Exists(unrelatedFolder));
            Assert.IsFalse(Directory.Exists(oldStagingFolder));
            Assert.IsTrue(Directory.Exists(recentStagingFolder));
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
        var renderRoot = Path.Combine(testRoot, "renders");
        var stagingRoot = Path.Combine(testRoot, "sources");
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                renderRoot,
                TimeProvider.System,
                detector,
                applicationFactory,
                launcher,
                FastWarmStartOptions,
                stagingRoot: stagingRoot);

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
            Assert.IsTrue(IsStrictDescendant(application.Presentations.OpenedPath, stagingRoot));
            Assert.IsFalse(IsStrictDescendant(application.Presentations.OpenedPath, renderRoot));
            Assert.IsFalse(File.Exists(application.Presentations.OpenedPath));
            Assert.IsFalse(Directory
                .EnumerateFiles(renderRoot, "*.ppt*", SearchOption.AllDirectories)
                .Any());
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
    public async Task BulkExportWithGapAndExtraFileFallsBackToExactSlideIndexes()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        await File.WriteAllBytesAsync(
            presentationPath,
            [1, 2, 3, 4],
            TestContext.CancellationToken);
        var application = new FakePowerPointApplication(
            userPresentationCount: 0,
            previewSlideCount: 3,
            bulkExportFileNames: ["Slide1.png", "Slide3.png", "Slide4.png"]);
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                new StubPowerPointProcessDetector(isRunning: true),
                new StubPowerPointApplicationFactory(application),
                new StubPowerPointProcessLauncher(),
                FastWarmStartOptions);

            var result = await renderer.RenderAsync(
                presentationPath,
                3,
                TestContext.CancellationToken);

            Assert.HasCount(3, result.SlideImages);
            Assert.AreEqual(1, application.Presentations.PreviewPresentation.ExportCount);
            Assert.AreEqual(3, application.Presentations.PreviewPresentation.Slides.IndividualExportCount);
            for (var slideNumber = 1; slideNumber <= 3; slideNumber++)
            {
                Assert.IsTrue(result.SlideImages.TryGetValue(slideNumber, out var imagePath));
                Assert.AreEqual($"slide-{slideNumber:D4}.png", Path.GetFileName(imagePath));
            }
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task ReadOnlySourceIsCleanedFromStagingWithoutChangingOriginalAttributes()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        var stagingRoot = Path.Combine(testRoot, "sources");
        await File.WriteAllBytesAsync(
            presentationPath,
            [1, 2, 3, 4],
            TestContext.CancellationToken);
        File.SetAttributes(presentationPath, FileAttributes.ReadOnly);
        var application = new FakePowerPointApplication(userPresentationCount: 0);
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                new StubPowerPointProcessDetector(isRunning: true),
                new StubPowerPointApplicationFactory(application),
                new StubPowerPointProcessLauncher(),
                FastWarmStartOptions,
                stagingRoot: stagingRoot);

            var result = await renderer.RenderAsync(
                presentationPath,
                1,
                TestContext.CancellationToken);

            Assert.HasCount(1, result.SlideImages);
            Assert.IsFalse(File.Exists(application.Presentations.OpenedPath));
            Assert.IsTrue((File.GetAttributes(presentationPath) & FileAttributes.ReadOnly) != 0);
            Assert.IsFalse(Directory
                .EnumerateFiles(stagingRoot, "*.ppt*", SearchOption.AllDirectories)
                .Any());
        }
        finally
        {
            if (File.Exists(presentationPath))
            {
                File.SetAttributes(presentationPath, FileAttributes.Normal);
            }

            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task FailedSourceCleanupStaysOutsidePreviewCacheAndRetriesOnDispose()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        var renderRoot = Path.Combine(testRoot, "renders");
        var stagingRoot = Path.Combine(testRoot, "sources");
        await File.WriteAllBytesAsync(
            presentationPath,
            [1, 2, 3, 4],
            TestContext.CancellationToken);
        var application = new FakePowerPointApplication(userPresentationCount: 0);
        PowerPointPresentationRenderer? renderer = null;
        var allowCleanup = 0;
        var cleanupAttempts = 0;
        try
        {
            renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                renderRoot,
                TimeProvider.System,
                new StubPowerPointProcessDetector(isRunning: true),
                new StubPowerPointApplicationFactory(application),
                new StubPowerPointProcessLauncher(),
                FastWarmStartOptions,
                stagingRoot: stagingRoot,
                stagingFolderDeletionOverride: folder =>
                {
                    Interlocked.Increment(ref cleanupAttempts);
                    if (Volatile.Read(ref allowCleanup) == 0)
                    {
                        return false;
                    }

                    Directory.Delete(folder, recursive: true);
                    return true;
                });

            var result = await renderer.RenderAsync(
                presentationPath,
                1,
                TestContext.CancellationToken);
            var stagedPath = application.Presentations.OpenedPath;

            Assert.HasCount(1, result.SlideImages);
            Assert.IsTrue(File.Exists(stagedPath));
            Assert.IsTrue(IsStrictDescendant(stagedPath, stagingRoot));
            Assert.IsFalse(IsStrictDescendant(stagedPath, renderRoot));
            Assert.IsFalse(Directory
                .EnumerateFiles(renderRoot, "*.ppt*", SearchOption.AllDirectories)
                .Any());
            Assert.AreEqual(1, Volatile.Read(ref cleanupAttempts));

            Volatile.Write(ref allowCleanup, 1);
            renderer.Dispose();
            renderer = null;

            Assert.IsFalse(File.Exists(stagedPath));
            Assert.AreEqual(2, Volatile.Read(ref cleanupAttempts));
        }
        finally
        {
            Volatile.Write(ref allowCleanup, 1);
            renderer?.Dispose();
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public void RegisteredLauncherFindsPerUserAppPathBeforeMachineRegistration()
    {
        var testRoot = CreateTestDirectory();
        var executablePath = Path.Combine(testRoot, "POWERPNT.EXE");
        File.WriteAllBytes(executablePath, [1, 2, 3, 4]);
        var requests = new List<(RegistryHive Hive, RegistryView View)>();
        try
        {
            var launcher = new RegisteredPowerPointProcessLauncher((hive, view) =>
            {
                requests.Add((hive, view));
                return (hive, view) == (RegistryHive.CurrentUser, RegistryView.Registry32)
                    ? $"\"{executablePath}\""
                    : null;
            });

            var registeredPath = launcher.FindRegisteredPowerPointExecutable();

            Assert.AreEqual(Path.GetFullPath(executablePath), registeredPath);
            Assert.HasCount(2, requests);
            Assert.AreEqual(
                (RegistryHive.CurrentUser, RegistryView.Registry64),
                requests[0]);
            Assert.AreEqual(
                (RegistryHive.CurrentUser, RegistryView.Registry32),
                requests[1]);
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public void RegisteredLauncherContinuesAfterInaccessiblePerUserRegistration()
    {
        var testRoot = CreateTestDirectory();
        var executablePath = Path.Combine(testRoot, "POWERPNT.EXE");
        File.WriteAllBytes(executablePath, [1, 2, 3, 4]);
        var requests = new List<(RegistryHive Hive, RegistryView View)>();
        try
        {
            var launcher = new RegisteredPowerPointProcessLauncher((hive, view) =>
            {
                requests.Add((hive, view));
                if (hive == RegistryHive.CurrentUser)
                {
                    throw new SecurityException("Per-user registration is inaccessible.");
                }

                return view == RegistryView.Registry64 ? executablePath : null;
            });

            var registeredPath = launcher.FindRegisteredPowerPointExecutable();

            Assert.AreEqual(Path.GetFullPath(executablePath), registeredPath);
            CollectionAssert.AreEqual(
                new[]
                {
                    (RegistryHive.CurrentUser, RegistryView.Registry64),
                    (RegistryHive.CurrentUser, RegistryView.Registry32),
                    (RegistryHive.LocalMachine, RegistryView.Registry64)
                },
                requests);
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task ManagedReparsePointRootFailsClosedBeforeWritingPreviewData()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        var renderRoot = Path.Combine(testRoot, "renders");
        await File.WriteAllBytesAsync(
            presentationPath,
            [1, 2, 3, 4],
            TestContext.CancellationToken);
        var application = new FakePowerPointApplication(userPresentationCount: 0);
        var applicationFactory = new StubPowerPointApplicationFactory(application);
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                renderRoot,
                TimeProvider.System,
                new StubPowerPointProcessDetector(isRunning: true),
                applicationFactory,
                new StubPowerPointProcessLauncher(),
                FastWarmStartOptions,
                fileAttributesReader: path => string.Equals(
                    Path.TrimEndingDirectorySeparator(Path.GetFullPath(path)),
                    Path.TrimEndingDirectorySeparator(Path.GetFullPath(renderRoot)),
                    StringComparison.OrdinalIgnoreCase)
                    ? FileAttributes.Directory | FileAttributes.ReparsePoint
                    : File.GetAttributes(path));

            var result = await renderer.RenderAsync(
                presentationPath,
                1,
                TestContext.CancellationToken);

            Assert.IsEmpty(result.SlideImages);
            Assert.AreEqual(0, applicationFactory.CreateCount);
            Assert.IsFalse(Directory.Exists(renderRoot));
        }
        finally
        {
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task LateWorkerCleanupAfterDisposeRetriesStagingFolder()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        var stagingRoot = Path.Combine(testRoot, "sources");
        await File.WriteAllBytesAsync(
            presentationPath,
            [1, 2, 3, 4],
            TestContext.CancellationToken);
        var application = new FakePowerPointApplication(userPresentationCount: 1);
        using var applicationFactory = new BlockingFirstPowerPointApplicationFactory(application);
        var cleanupCompleted = new TaskCompletionSource<object?>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var cleanupAttempts = 0;
        PowerPointPresentationRenderer? renderer = null;
        try
        {
            renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                new StubPowerPointProcessDetector(isRunning: true),
                applicationFactory,
                new StubPowerPointProcessLauncher(),
                FastWarmStartOptions,
                TimeSpan.FromMilliseconds(250),
                stagingRoot,
                folder =>
                {
                    if (Interlocked.Increment(ref cleanupAttempts) == 1)
                    {
                        return false;
                    }

                    Directory.Delete(folder, recursive: true);
                    cleanupCompleted.TrySetResult(null);
                    return true;
                });

            var timedOutRender = renderer.RenderAsync(
                presentationPath,
                1,
                TestContext.CancellationToken);
            await applicationFactory.FirstActivationStarted.WaitAsync(
                TimeSpan.FromSeconds(2),
                TestContext.CancellationToken);
            var result = await timedOutRender.WaitAsync(
                TimeSpan.FromSeconds(2),
                TestContext.CancellationToken);
            StringAssert.Contains(result.Status, "timed out");

            renderer.Dispose();
            renderer = null;
            applicationFactory.ReleaseFirstActivation();

            await cleanupCompleted.Task.WaitAsync(
                TimeSpan.FromSeconds(2),
                TestContext.CancellationToken);
            Assert.AreEqual(2, Volatile.Read(ref cleanupAttempts));
            Assert.IsFalse(Directory
                .EnumerateFiles(stagingRoot, "*.ppt*", SearchOption.AllDirectories)
                .Any());
        }
        finally
        {
            renderer?.Dispose();
            applicationFactory.ReleaseFirstActivation();
            await applicationFactory.FirstActivationReturned.WaitAsync(TimeSpan.FromSeconds(2));
            DeleteTestDirectory(testRoot);
        }
    }

    [TestMethod]
    public async Task TimedOutWorkerRetainsRenderLeaseUntilItsStaThreadExits()
    {
        var testRoot = CreateTestDirectory();
        var presentationPath = Path.Combine(testRoot, "source.pptx");
        await File.WriteAllBytesAsync(
            presentationPath,
            [1, 2, 3, 4],
            TestContext.CancellationToken);
        var application = new FakePowerPointApplication(userPresentationCount: 1);
        using var applicationFactory = new BlockingFirstPowerPointApplicationFactory(application);
        try
        {
            using var renderer = new PowerPointPresentationRenderer(
                new RecordingDiagnostics(),
                Path.Combine(testRoot, "renders"),
                TimeProvider.System,
                new StubPowerPointProcessDetector(isRunning: true),
                applicationFactory,
                new StubPowerPointProcessLauncher(),
                FastWarmStartOptions,
                TimeSpan.FromMilliseconds(250));

            var timedOutRender = renderer.RenderAsync(
                presentationPath,
                1,
                TestContext.CancellationToken);
            await applicationFactory.FirstActivationStarted.WaitAsync(
                TimeSpan.FromSeconds(2),
                TestContext.CancellationToken);

            var timedOutResult = await timedOutRender.WaitAsync(
                TimeSpan.FromSeconds(2),
                TestContext.CancellationToken);

            StringAssert.Contains(timedOutResult.Status, "timed out");
            Assert.AreEqual(1, applicationFactory.CreateCount);

            using var queuedCancellation = new CancellationTokenSource();
            var queuedRender = renderer.RenderAsync(
                presentationPath,
                1,
                queuedCancellation.Token);

            Assert.IsFalse(queuedRender.IsCompleted);
            Assert.AreEqual(1, applicationFactory.CreateCount);
            await queuedCancellation.CancelAsync();
            try
            {
                await queuedRender;
                Assert.Fail("A render cancelled while waiting for the active worker should not run.");
            }
            catch (OperationCanceledException) when (queuedCancellation.IsCancellationRequested)
            {
                // Expected: callers can cancel while the timed-out worker retains the lease.
            }

            Assert.AreEqual(1, applicationFactory.CreateCount);
            applicationFactory.ReleaseFirstActivation();

            var subsequentResult = await renderer.RenderAsync(
                    presentationPath,
                    1,
                    TestContext.CancellationToken)
                .WaitAsync(TimeSpan.FromSeconds(2), TestContext.CancellationToken);

            Assert.AreEqual(2, applicationFactory.CreateCount);
            Assert.AreEqual(1, applicationFactory.MaximumConcurrentActivations);
            Assert.AreEqual(1, subsequentResult.SlideImages.Count);
        }
        finally
        {
            applicationFactory.ReleaseFirstActivation();
            await applicationFactory.FirstActivationReturned.WaitAsync(TimeSpan.FromSeconds(2));
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
            StringAssert.Contains(summary, "Version: 0.3.0");
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

    private static bool IsStrictDescendant(string candidatePath, string rootPath)
    {
        var fullCandidate = Path.TrimEndingDirectorySeparator(Path.GetFullPath(candidatePath));
        var fullRoot = Path.TrimEndingDirectorySeparator(Path.GetFullPath(rootPath));
        return fullCandidate.StartsWith(
            fullRoot + Path.DirectorySeparatorChar,
            StringComparison.OrdinalIgnoreCase);
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

    private sealed class BlockingFirstPowerPointApplicationFactory(
        FakePowerPointApplication application) : IPowerPointApplicationFactory, IDisposable
    {
        private readonly ManualResetEventSlim _releaseFirstActivation = new(initialState: false);
        private readonly TaskCompletionSource<object?> _firstActivationReturned = new(
            TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _firstActivationStarted = new(
            TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeActivations;
        private int _createCount;
        private int _maximumConcurrentActivations;

        public int CreateCount => Volatile.Read(ref _createCount);

        public Task FirstActivationReturned => _firstActivationReturned.Task;

        public Task FirstActivationStarted => _firstActivationStarted.Task;

        public int MaximumConcurrentActivations =>
            Volatile.Read(ref _maximumConcurrentActivations);

        public bool IsPowerPointInstalled() => true;

        public object CreateApplication()
        {
            var activationNumber = Interlocked.Increment(ref _createCount);
            var activeActivations = Interlocked.Increment(ref _activeActivations);
            SetMaximumConcurrentActivations(activeActivations);
            try
            {
                if (activationNumber == 1)
                {
                    _firstActivationStarted.TrySetResult(null);
                    _releaseFirstActivation.Wait();
                }

                return application;
            }
            finally
            {
                Interlocked.Decrement(ref _activeActivations);
                if (activationNumber == 1)
                {
                    _firstActivationReturned.TrySetResult(null);
                }
            }
        }

        public void ReleaseFirstActivation() => _releaseFirstActivation.Set();

        public void Dispose() => _releaseFirstActivation.Dispose();

        private void SetMaximumConcurrentActivations(int candidate)
        {
            var current = Volatile.Read(ref _maximumConcurrentActivations);
            while (candidate > current)
            {
                var observed = Interlocked.CompareExchange(
                    ref _maximumConcurrentActivations,
                    candidate,
                    current);
                if (observed == current)
                {
                    return;
                }

                current = observed;
            }
        }
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

    public sealed class FakePowerPointApplication(
        int userPresentationCount,
        Action? onQuit = null,
        int previewSlideCount = 1,
        IReadOnlyList<string>? bulkExportFileNames = null)
    {
        public int AutomationSecurity { get; set; } = 2;

        public FakePowerPointPresentations Presentations { get; } = new(
            userPresentationCount,
            previewSlideCount,
            bulkExportFileNames);

        public bool QuitCalled { get; private set; }

        public void Quit()
        {
            QuitCalled = true;
            onQuit?.Invoke();
        }
    }

    public sealed class FakePowerPointPresentations(
        int userPresentationCount,
        int previewSlideCount,
        IReadOnlyList<string>? bulkExportFileNames)
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
            PreviewPresentation = new FakePowerPointPresentation(
                this,
                previewSlideCount,
                bulkExportFileNames);
            return PreviewPresentation;
        }

        public void ClosePreview()
        {
            _previewPresentationOpen = false;
            PreviewPresentationClosed = true;
        }
    }

    public sealed class FakePowerPointPresentation(
        FakePowerPointPresentations owner,
        int slideCount,
        IReadOnlyList<string>? bulkExportFileNames)
    {
        public FakePowerPointSlides Slides { get; } = new(slideCount);

        public FakePowerPointPageSetup PageSetup { get; } = new();

        public int ExportCount { get; private set; }

        public void Export(string folder, string format, int width, int height)
        {
            ExportCount++;
            Directory.CreateDirectory(folder);
            var fileNames = bulkExportFileNames ?? Enumerable
                .Range(1, slideCount)
                .Select(index => $"Slide{index}.png")
                .ToArray();
            foreach (var fileName in fileNames)
            {
                File.WriteAllBytes(Path.Combine(folder, fileName), [1, 2, 3, 4]);
            }
        }

        public void Close() => owner.ClosePreview();
    }

    public sealed class FakePowerPointSlides(int count)
    {
        public int Count { get; } = count;

        public int IndividualExportCount { get; private set; }

        public FakePowerPointSlide Item(int index)
        {
            if (index < 1 || index > Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return new FakePowerPointSlide(this);
        }

        public void RecordIndividualExport() => IndividualExportCount++;
    }

    public sealed class FakePowerPointSlide(FakePowerPointSlides owner)
    {
        public void Export(string path, string format, int width, int height)
        {
            owner.RecordIndividualExport();
            File.WriteAllBytes(path, [1, 2, 3, 4]);
        }
    }

    public sealed class FakePowerPointPageSetup
    {
        public double SlideWidth { get; } = 960;

        public double SlideHeight { get; } = 540;
    }

    public TestContext TestContext { get; set; } = null!;
}
