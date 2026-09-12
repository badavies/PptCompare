using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Security;
using Microsoft.Win32;

namespace PptCompare.Services;

public sealed class PowerPointPresentationRenderer : IPresentationRenderer
{
    private const int ComServerExecutionFailed = unchecked((int)0x80080005);
    private const int MaxCachedPresentations = 8;
    private const long MaxRenderedImageBytes = 50L * 1024 * 1024;
    private const long MaxRenderedPresentationBytes = 500L * 1024 * 1024;
    private static readonly TimeSpan RenderTimeout = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan StaleRenderFolderAge = TimeSpan.FromHours(24);
    private readonly Lock _renderFolderGate = new();
    private readonly SemaphoreSlim _renderSemaphore = new(1, 1);
    private readonly Queue<string> _renderFolders = new();
    private readonly IPowerPointApplicationFactory _applicationFactory;
    private readonly IApplicationDiagnostics _diagnostics;
    private readonly IPowerPointProcessLauncher _powerPointProcessLauncher;
    private readonly IPowerPointProcessDetector _powerPointProcessDetector;
    private readonly string _renderRoot;
    private readonly TimeProvider _timeProvider;
    private readonly PowerPointWarmStartOptions _warmStartOptions;
    private bool _disposed;

    public PowerPointPresentationRenderer(IApplicationDiagnostics diagnostics)
        : this(
            diagnostics,
            Path.Combine(Path.GetTempPath(), "PptCompare", "renders"),
            TimeProvider.System,
            new SystemPowerPointProcessDetector(),
            new ComPowerPointApplicationFactory(),
            new RegisteredPowerPointProcessLauncher(),
            PowerPointWarmStartOptions.Default)
    {
    }

    internal PowerPointPresentationRenderer(
        IApplicationDiagnostics diagnostics,
        string renderRoot,
        TimeProvider timeProvider,
        IPowerPointProcessDetector? powerPointProcessDetector = null,
        IPowerPointApplicationFactory? applicationFactory = null,
        IPowerPointProcessLauncher? powerPointProcessLauncher = null,
        PowerPointWarmStartOptions? warmStartOptions = null)
    {
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        ArgumentException.ThrowIfNullOrWhiteSpace(renderRoot);
        _renderRoot = Path.GetFullPath(renderRoot);
        _timeProvider = timeProvider ?? throw new ArgumentNullException(nameof(timeProvider));
        _powerPointProcessDetector = powerPointProcessDetector ?? new SystemPowerPointProcessDetector();
        _applicationFactory = applicationFactory ?? new ComPowerPointApplicationFactory();
        _powerPointProcessLauncher = powerPointProcessLauncher ?? new RegisteredPowerPointProcessLauncher();
        _warmStartOptions = warmStartOptions ?? PowerPointWarmStartOptions.Default;
        _warmStartOptions.Validate();
        CleanupStaleRenderFolders();
    }

    public async Task<SlideRenderingResult> RenderAsync(
        string presentationPath,
        int slideCount,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentException.ThrowIfNullOrWhiteSpace(presentationPath);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(slideCount);

        await _renderSemaphore.WaitAsync(cancellationToken);
        try
        {
            var trace = new PowerPointRenderTrace(_diagnostics);
            trace.Record("request", $"slides={slideCount.ToString(CultureInfo.InvariantCulture)}");
            bool isPowerPointInstalled;
            try
            {
                trace.Record("installation-check-started");
                isPowerPointInstalled = _applicationFactory.IsPowerPointInstalled();
                trace.Record(
                    "installation-check-completed",
                    $"installed={isPowerPointInstalled.ToString(CultureInfo.InvariantCulture)}");
            }
            catch (Exception exception)
            {
                trace.RecordFailure("installation-check", exception);
                _diagnostics.RecordException("PowerPointInstallationCheckFailed", exception);
                return new SlideRenderingResult(
                    new Dictionary<int, string>(),
                    "Microsoft PowerPoint could not be inspected; using the built-in slide preview.");
            }

            if (!isPowerPointInstalled)
            {
                return new SlideRenderingResult(
                    new Dictionary<int, string>(),
                    "Microsoft PowerPoint was not found; using the built-in slide preview.");
            }

            var powerPointWasAlreadyRunning = DetectExistingPowerPointSession();
            trace.Record(
                "process-check-completed",
                $"alreadyRunning={powerPointWasAlreadyRunning.ToString(CultureInfo.InvariantCulture)}");

            using var renderCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken);
            var renderCancellationToken = renderCancellation.Token;
            var completion = new TaskCompletionSource<SlideRenderingResult>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
                RenderOnStaThread(
                    presentationPath,
                    slideCount,
                    powerPointWasAlreadyRunning,
                    trace,
                    completion,
                    renderCancellationToken))
            {
                IsBackground = true,
                Name = "PptCompare PowerPoint renderer"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            try
            {
                return await completion.Task.WaitAsync(RenderTimeout, cancellationToken);
            }
            catch (TimeoutException)
            {
                await renderCancellation.CancelAsync();
                Debug.WriteLine(
                    $"PowerPoint rendering timed out after {RenderTimeout.TotalSeconds:N0} seconds.");
                _diagnostics.RecordEvent("PreviewRenderTimedOut");
                trace.Record("timed-out");
                return new SlideRenderingResult(
                    new Dictionary<int, string>(),
                    "PowerPoint rendering timed out; using the built-in slide preview.");
            }
        }
        finally
        {
            _renderSemaphore.Release();
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "Optional Office automation must fail closed and return the built-in renderer fallback.")]
    private void RenderOnStaThread(
        string presentationPath,
        int expectedSlideCount,
        bool powerPointWasAlreadyRunning,
        PowerPointRenderTrace trace,
        TaskCompletionSource<SlideRenderingResult> completion,
        CancellationToken cancellationToken)
    {
        object? application = null;
        object? presentations = null;
        object? presentation = null;
        object? slides = null;
        object? pageSetup = null;
        string? outputFolder = null;
        string? stagedPresentationPath = null;
        SlideRenderingResult? renderingResult = null;
        var operationCancelled = false;
        var ownedPresentationClosed = true;
        int? originalAutomationSecurity = null;
        var automationSecurityNeedsRestoring = false;
        var stage = "initialization";

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            stage = "temporary-copy";
            trace.Record("temporary-copy-started");
            outputFolder = CreateRenderFolder();
            stagedPresentationPath = CreateStagedPresentationCopy(
                presentationPath,
                outputFolder);
            trace.Record("temporary-copy-completed");

            if (!powerPointWasAlreadyRunning)
            {
                stage = "application-warm-start";
                WarmStartPowerPoint(trace, cancellationToken);
            }

            stage = "application-activation";
            application = CreatePowerPointApplicationWithRetry(trace, cancellationToken);

            stage = "automation-security";
            dynamic powerPoint = application;
            originalAutomationSecurity = Convert.ToInt32(
                powerPoint.AutomationSecurity,
                CultureInfo.InvariantCulture);
            powerPoint.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            automationSecurityNeedsRestoring = true;
            trace.Record("automation-security-configured");
            presentations = powerPoint.Presentations;
            dynamic presentationCollection = presentations;
            stage = "presentation-open";
            trace.Record("presentation-open-started");
            try
            {
                presentation = presentationCollection.Open(
                    stagedPresentationPath,
                    -1, // ReadOnly: msoTrue
                    0,  // Untitled: msoFalse
                    0); // WithWindow: msoFalse
            }
            finally
            {
                automationSecurityNeedsRestoring = !TryRestoreAutomationSecurity(
                    application,
                    originalAutomationSecurity.Value);
                trace.Record(
                    "automation-security-restored",
                    $"restored={(!automationSecurityNeedsRestoring).ToString(CultureInfo.InvariantCulture)}");
            }
            trace.Record("presentation-open-completed");

            stage = "presentation-inspection";
            dynamic openedPresentation = presentation;
            slides = openedPresentation.Slides;
            dynamic slideCollection = slides;
            var count = Math.Min((int)slideCollection.Count, expectedSlideCount);
            pageSetup = openedPresentation.PageSetup;
            dynamic presentationPageSetup = pageSetup;
            var exportSize = CalculateExportSize(
                (double)presentationPageSetup.SlideWidth,
                (double)presentationPageSetup.SlideHeight);

            cancellationToken.ThrowIfCancellationRequested();
            var bulkFolder = Path.Combine(outputFolder, "bulk");
            Directory.CreateDirectory(bulkFolder);
            Dictionary<int, string> images;
            stage = "bulk-export";
            trace.Record("bulk-export-started");
            try
            {
                openedPresentation.Export(
                    bulkFolder,
                    "PNG",
                    exportSize.Width,
                    exportSize.Height);
                cancellationToken.ThrowIfCancellationRequested();
                images = CollectExportedImages(bulkFolder, count);
                trace.Record(
                    "bulk-export-completed",
                    $"images={images.Count.ToString(CultureInfo.InvariantCulture)}");
            }
            catch (Exception exception) when (
                exception is COMException or IOException or UnauthorizedAccessException)
            {
                Debug.WriteLine($"Whole-presentation PowerPoint export failed: {exception}");
                _diagnostics.RecordException("PreviewBulkExportFailed", exception);
                trace.RecordFailure("bulk-export", exception);
                images = [];
            }

            if (images.Count != count)
            {
                // Some older or repaired decks fail whole-presentation export even though
                // PowerPoint can render their slides individually.
                DeleteRenderFolder(bulkFolder);
                stage = "individual-export";
                trace.Record("individual-export-started");
                images = ExportSlidesIndividually(
                    slideCollection,
                    outputFolder,
                    count,
                    exportSize.Width,
                    exportSize.Height,
                    cancellationToken);
                trace.Record(
                    "individual-export-completed",
                    $"images={images.Count.ToString(CultureInfo.InvariantCulture)}");
            }

            stage = "render-completed";
            renderingResult = new SlideRenderingResult(
                images,
                images.Count == count
                    ? "High-fidelity slide previews rendered by Microsoft PowerPoint."
                    : "Some slides could not be rendered; using built-in previews where necessary.");
            trace.Record(
                stage,
                $"images={images.Count.ToString(CultureInfo.InvariantCulture)}");
        }
        catch (OperationCanceledException)
        {
            operationCancelled = true;
            trace.Record("cancelled", $"lastStage={stage}");
        }
        catch (Exception exception)
        {
            Debug.WriteLine($"PowerPoint rendering failed: {exception}");
            _diagnostics.RecordException("PreviewRenderFailed", exception);
            trace.RecordFailure(stage, exception);
            if (string.Equals(stage, "application-activation", StringComparison.Ordinal))
            {
                RecordProcessStateAfterActivationFailure(trace);
            }

            renderingResult = new SlideRenderingResult(
                new Dictionary<int, string>(),
                GetRenderingFailureStatus(exception));
        }
        finally
        {
            if (automationSecurityNeedsRestoring &&
                originalAutomationSecurity is { } security)
            {
                TryRestoreAutomationSecurity(application, security);
            }

            if (presentation is not null)
            {
                trace.Record("presentation-close-started");
                ownedPresentationClosed = TryCloseOwnedPresentation(presentation);
                trace.Record(
                    "presentation-close-completed",
                    $"closed={ownedPresentationClosed.ToString(CultureInfo.InvariantCulture)}");
            }

            var canQuitApplication = CanQuitPowerPointWithoutUserDataLoss(
                presentations,
                powerPointWasAlreadyRunning,
                ownedPresentationClosed);
            trace.Record(
                "application-shutdown-policy",
                $"preExisting={powerPointWasAlreadyRunning.ToString(CultureInfo.InvariantCulture)};quit={canQuitApplication.ToString(CultureInfo.InvariantCulture)}");
            ReleaseComObject(slides);
            ReleaseComObject(pageSetup);
            ReleaseComObject(presentation);
            ReleaseComObject(presentations);

            if (canQuitApplication)
            {
                try
                {
                    trace.Record("application-quit-started");
                    ((dynamic)application!).Quit();
                    trace.Record("application-quit-completed");
                    WaitForPowerPointProcessExit(trace);
                }
                catch (COMException exception)
                {
                    _diagnostics.RecordException("PowerPointAutomationShutdownFailed", exception);
                    trace.RecordFailure("application-quit", exception);
                }
            }

            ReleaseComObject(application);
            DeleteStagedPresentation(stagedPresentationPath);
            trace.Record("temporary-copy-cleanup-requested");

            if (!operationCancelled &&
                outputFolder is not null &&
                renderingResult is { SlideImages.Count: > 0 })
            {
                RememberRenderFolder(outputFolder);
                outputFolder = null;
            }

            if (outputFolder is not null)
            {
                DeleteRenderFolder(outputFolder);
            }

            if (operationCancelled)
            {
                completion.TrySetCanceled(cancellationToken);
            }
            else
            {
                trace.Record("cleanup-completed");
                completion.TrySetResult(renderingResult ?? new SlideRenderingResult(
                    new Dictionary<int, string>(),
                    "PowerPoint rendering was unavailable; using the built-in slide preview."));
            }
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "Failure to inspect running processes must assume a shared session so it can never be terminated.")]
    private bool DetectExistingPowerPointSession()
    {
        try
        {
            var isRunning = _powerPointProcessDetector.IsPowerPointRunning();
            if (isRunning)
            {
                _diagnostics.RecordEvent("PreviewRenderingUsingExistingPowerPoint");
            }

            return isRunning;
        }
        catch (Exception exception)
        {
            _diagnostics.RecordException("PowerPointProcessInspectionFailed", exception);
            return true;
        }
    }

    private void WarmStartPowerPoint(
        PowerPointRenderTrace trace,
        CancellationToken cancellationToken)
    {
        trace.Record("warm-start-launch-started");
        _powerPointProcessLauncher.StartPowerPoint();
        trace.Record("warm-start-launch-requested");

        var wait = Stopwatch.StartNew();
        while (wait.Elapsed < _warmStartOptions.ProcessStartTimeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_powerPointProcessDetector.IsPowerPointRunning())
            {
                trace.Record(
                    "warm-start-process-detected",
                    $"waitMs={wait.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)}");
                WaitWithCancellation(
                    _warmStartOptions.AutomationReadyDelay,
                    cancellationToken);
                trace.Record("warm-start-ready-delay-completed");
                return;
            }

            WaitWithCancellation(_warmStartOptions.ProcessPollInterval, cancellationToken);
        }

        throw new TimeoutException("PowerPoint did not start within the permitted time.");
    }

    private object CreatePowerPointApplicationWithRetry(
        PowerPointRenderTrace trace,
        CancellationToken cancellationToken)
    {
        for (var attempt = 1; attempt <= _warmStartOptions.MaxActivationAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            trace.Record(
                "application-activation-started",
                $"attempt={attempt.ToString(CultureInfo.InvariantCulture)}");
            try
            {
                var application = _applicationFactory.CreateApplication();
                trace.Record(
                    "application-activation-completed",
                    $"attempt={attempt.ToString(CultureInfo.InvariantCulture)}");
                return application;
            }
            catch (COMException exception) when (
                exception.HResult == ComServerExecutionFailed &&
                attempt < _warmStartOptions.MaxActivationAttempts)
            {
                trace.RecordFailure($"application-activation-attempt-{attempt}", exception);
                if (!_powerPointProcessDetector.IsPowerPointRunning())
                {
                    throw;
                }

                WaitWithCancellation(_warmStartOptions.ActivationRetryDelay, cancellationToken);
            }
        }

        throw new InvalidOperationException("PowerPoint automation activation did not complete.");
    }

    private void WaitForPowerPointProcessExit(PowerPointRenderTrace trace)
    {
        var wait = Stopwatch.StartNew();
        while (wait.Elapsed < _warmStartOptions.ProcessExitTimeout)
        {
            try
            {
                if (!_powerPointProcessDetector.IsPowerPointRunning())
                {
                    trace.Record(
                        "application-process-exit-detected",
                        $"waitMs={wait.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)}");
                    return;
                }
            }
            catch (Exception exception)
            {
                trace.RecordFailure("application-process-exit-check", exception);
                return;
            }

            WaitWithCancellation(
                _warmStartOptions.ProcessPollInterval,
                CancellationToken.None);
        }

        trace.Record("application-process-exit-wait-expired");
    }

    private static void WaitWithCancellation(
        TimeSpan delay,
        CancellationToken cancellationToken)
    {
        if (delay <= TimeSpan.Zero)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return;
        }

        if (cancellationToken.WaitHandle.WaitOne(delay))
        {
            cancellationToken.ThrowIfCancellationRequested();
        }
    }

    private void RecordProcessStateAfterActivationFailure(PowerPointRenderTrace trace)
    {
        if (!trace.IsEnabled)
        {
            return;
        }

        try
        {
            var isRunning = _powerPointProcessDetector.IsPowerPointRunning();
            trace.Record(
                "activation-failure-process-check",
                $"running={isRunning.ToString(CultureInfo.InvariantCulture)}");
        }
        catch (Exception exception)
        {
            trace.RecordFailure("activation-failure-process-check", exception);
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "Failure to restore an application-wide setting is logged and retried during cleanup.")]
    private bool TryRestoreAutomationSecurity(object? application, int originalValue)
    {
        if (application is null)
        {
            return false;
        }

        try
        {
            ((dynamic)application).AutomationSecurity = originalValue;
            return true;
        }
        catch (Exception exception)
        {
            _diagnostics.RecordException("PowerPointAutomationSecurityRestoreFailed", exception);
            return false;
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "A failed close must leave the shared PowerPoint application running and return the built-in fallback safely.")]
    private bool TryCloseOwnedPresentation(object presentation)
    {
        try
        {
            ((dynamic)presentation).Close();
            return true;
        }
        catch (Exception exception)
        {
            _diagnostics.RecordException("PowerPointPreviewCloseFailed", exception);
            return false;
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "An uncertain PowerPoint state must fail closed and leave the application running.")]
    private bool CanQuitPowerPointWithoutUserDataLoss(
        object? presentationCollection,
        bool powerPointWasAlreadyRunning,
        bool ownedPresentationClosed)
    {
        if (powerPointWasAlreadyRunning)
        {
            _diagnostics.RecordEvent(
                "PowerPointAutomationShutdownSkipped",
                "reason=pre-existing-session");
            return false;
        }

        if (!ownedPresentationClosed)
        {
            _diagnostics.RecordEvent(
                "PowerPointAutomationShutdownSkipped",
                "reason=preview-close-failed");
            return false;
        }

        if (presentationCollection is null)
        {
            return false;
        }

        try
        {
            dynamic presentations = presentationCollection;
            var openPresentationCount = (int)presentations.Count;
            if (openPresentationCount == 0)
            {
                return true;
            }

            _diagnostics.RecordEvent(
                "PowerPointAutomationShutdownSkipped",
                $"openPresentations={openPresentationCount.ToString(CultureInfo.InvariantCulture)}");
            return false;
        }
        catch (Exception exception)
        {
            _diagnostics.RecordException("PowerPointShutdownSafetyCheckFailed", exception);
            return false;
        }
    }

    private static string CreateStagedPresentationCopy(
        string presentationPath,
        string outputFolder)
    {
        var sourcePath = Path.GetFullPath(presentationPath);
        var extension = Path.GetExtension(sourcePath);
        if (!extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".pptm", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException(
                "Only .pptx and .pptm presentations can be rendered.");
        }

        var stagingFolder = Path.Combine(outputFolder, "source");
        Directory.CreateDirectory(stagingFolder);
        var stagedPath = Path.Combine(stagingFolder, $"preview{extension.ToLowerInvariant()}");
        File.Copy(sourcePath, stagedPath, overwrite: false);
        return stagedPath;
    }

    private void DeleteStagedPresentation(string? stagedPresentationPath)
    {
        if (stagedPresentationPath is null)
        {
            return;
        }

        var stagingFolder = Path.GetDirectoryName(stagedPresentationPath);
        if (stagingFolder is not null && !DeleteRenderFolder(stagingFolder))
        {
            _diagnostics.RecordEvent("PreviewSourceCleanupDeferred");
        }
    }

    private string CreateRenderFolder()
    {
        Directory.CreateDirectory(_renderRoot);
        var folder = Path.Combine(_renderRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static (int Width, int Height) CalculateExportSize(
        double slideWidth,
        double slideHeight)
    {
        const int longestEdge = 1600;
        if (!double.IsFinite(slideWidth) || !double.IsFinite(slideHeight) ||
            slideWidth <= 0 || slideHeight <= 0)
        {
            return (1600, 900);
        }

        return slideWidth >= slideHeight
            ? (longestEdge, Math.Max(1, (int)Math.Round(
                longestEdge * slideHeight / slideWidth,
                MidpointRounding.AwayFromZero)))
            : (Math.Max(1, (int)Math.Round(
                longestEdge * slideWidth / slideHeight,
                MidpointRounding.AwayFromZero)), longestEdge);
    }

    private static Dictionary<int, string> CollectExportedImages(
        string outputFolder,
        int expectedCount)
    {
        var files = Directory
            .EnumerateFiles(outputFolder, "*.png", SearchOption.TopDirectoryOnly)
            .Select(path => new FileInfo(path))
            .Where(file => file.Length is > 0 and <= MaxRenderedImageBytes)
            .OrderBy(file => ReadTrailingNumber(file.Name))
            .ThenBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
            .Take(expectedCount)
            .ToList();
        var images = new Dictionary<int, string>();
        long totalBytes = 0;

        for (var index = 0; index < files.Count; index++)
        {
            totalBytes += files[index].Length;
            if (totalBytes > MaxRenderedPresentationBytes)
            {
                break;
            }

            images[index + 1] = files[index].FullName;
        }

        return images;
    }

    private Dictionary<int, string> ExportSlidesIndividually(
        object slideCollection,
        string outputFolder,
        int slideCount,
        int width,
        int height,
        CancellationToken cancellationToken)
    {
        dynamic slides = slideCollection;
        var images = new Dictionary<int, string>();
        long totalBytes = 0;
        for (var index = 1; index <= slideCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            object? slide = null;
            try
            {
                slide = slides.Item(index);
                var imagePath = Path.Combine(outputFolder, $"slide-{index:D4}.png");
                ((dynamic)slide).Export(imagePath, "PNG", width, height);
                var image = new FileInfo(imagePath);
                if (!image.Exists || image.Length is <= 0 or > MaxRenderedImageBytes)
                {
                    continue;
                }

                totalBytes += image.Length;
                if (totalBytes > MaxRenderedPresentationBytes)
                {
                    break;
                }

                images[index] = image.FullName;
            }
            catch (COMException exception)
            {
                Debug.WriteLine($"PowerPoint could not render slide {index}: {exception}");
                _diagnostics.RecordException("PreviewSlideExportFailed", exception);
                // Keep rendering the remaining slides; the built-in preview covers this one.
            }
            catch (IOException exception)
            {
                Debug.WriteLine($"The rendered image for slide {index} could not be saved: {exception}");
                _diagnostics.RecordException("PreviewSlideFileFailed", exception);
            }
            catch (UnauthorizedAccessException exception)
            {
                Debug.WriteLine($"The rendered image for slide {index} could not be accessed: {exception}");
                _diagnostics.RecordException("PreviewSlideAccessFailed", exception);
            }
            finally
            {
                ReleaseComObject(slide);
            }
        }

        return images;
    }

    private static int ReadTrailingNumber(string fileName)
    {
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var end = stem.Length;
        var start = end;
        while (start > 0 && char.IsDigit(stem[start - 1]))
        {
            start--;
        }

        return start < end && int.TryParse(
            stem.AsSpan(start, end - start),
            NumberStyles.None,
            CultureInfo.InvariantCulture,
            out var number)
            ? number
            : int.MaxValue;
    }

    private static string GetRenderingFailureStatus(Exception exception) =>
        exception is COMException { HResult: ComServerExecutionFailed }
            ? "Microsoft PowerPoint could not start. Open PowerPoint once and respond to any recovery or safe-mode prompt, then try again. The built-in preview remains available."
            : "PowerPoint rendering was unavailable; using the built-in slide preview.";

    private void RememberRenderFolder(string folder)
    {
        var foldersToDelete = new List<string>();
        lock (_renderFolderGate)
        {
            if (_disposed)
            {
                foldersToDelete.Add(folder);
            }
            else
            {
                _renderFolders.Enqueue(folder);
                while (_renderFolders.Count > MaxCachedPresentations)
                {
                    foldersToDelete.Add(_renderFolders.Dequeue());
                }
            }
        }

        foreach (var folderToDelete in foldersToDelete)
        {
            DeleteRenderFolder(folderToDelete);
        }
    }

    private void CleanupStaleRenderFolders()
    {
        try
        {
            if (!Directory.Exists(_renderRoot))
            {
                return;
            }

            var cutoff = _timeProvider.GetUtcNow().UtcDateTime - StaleRenderFolderAge;
            var deletedCount = 0;
            foreach (var folder in Directory.EnumerateDirectories(
                         _renderRoot,
                         "*",
                         SearchOption.TopDirectoryOnly))
            {
                try
                {
                    var folderName = Path.GetFileName(folder);
                    var attributes = File.GetAttributes(folder);
                    if (!Guid.TryParseExact(folderName, "N", out _) ||
                        (attributes & FileAttributes.ReparsePoint) != 0 ||
                        Directory.GetLastWriteTimeUtc(folder) > cutoff)
                    {
                        continue;
                    }

                    if (DeleteRenderFolder(folder))
                    {
                        deletedCount++;
                    }
                }
                catch (Exception exception) when (
                    exception is IOException or UnauthorizedAccessException or SecurityException)
                {
                    _diagnostics.RecordException("StalePreviewFolderInspectionFailed", exception);
                }
            }

            if (deletedCount > 0)
            {
                _diagnostics.RecordEvent(
                    "StalePreviewCleanupCompleted",
                    $"folders={deletedCount.ToString(CultureInfo.InvariantCulture)}");
            }
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or SecurityException)
        {
            _diagnostics.RecordException("StalePreviewCleanupFailed", exception);
        }
    }

    private bool DeleteRenderFolder(string folder)
    {
        try
        {
            var fullPath = Path.GetFullPath(folder);
            var rootWithSeparator = _renderRoot.TrimEnd(Path.DirectorySeparatorChar) +
                                    Path.DirectorySeparatorChar;
            if (fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase) &&
                Directory.Exists(fullPath))
            {
                Directory.Delete(fullPath, true);
                return true;
            }
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or SecurityException)
        {
            _diagnostics.RecordException("PreviewFolderCleanupFailed", exception);
        }

        return false;
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }

    private sealed class PowerPointRenderTrace(IApplicationDiagnostics diagnostics)
    {
        private readonly IDetailedApplicationDiagnostics? _diagnostics =
            diagnostics as IDetailedApplicationDiagnostics;
        private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
        private readonly string _renderId = Guid.NewGuid().ToString("N")[..8];

        public bool IsEnabled => _diagnostics is { IsDebugLoggingEnabled: true };

        public void Record(string stage, string? nonSensitiveDetail = null)
        {
            if (_diagnostics is not { IsDebugLoggingEnabled: true } diagnostics)
            {
                return;
            }

            var detail = string.Create(
                CultureInfo.InvariantCulture,
                $"renderId={_renderId};stage={stage};elapsedMs={_stopwatch.ElapsedMilliseconds}");
            if (!string.IsNullOrWhiteSpace(nonSensitiveDetail))
            {
                detail += $";{nonSensitiveDetail}";
            }

            diagnostics.RecordDebugEvent("PowerPointPreviewTrace", detail);
        }

        public void RecordFailure(string stage, Exception exception)
        {
            ArgumentNullException.ThrowIfNull(exception);
            Record(
                "failure",
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"failedStage={stage};hresult=0x{exception.HResult:X8};type={exception.GetType().Name}"));
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        List<string> foldersToDelete;
        lock (_renderFolderGate)
        {
            _disposed = true;
            foldersToDelete = [.. _renderFolders];
            _renderFolders.Clear();
        }

        foreach (var folder in foldersToDelete)
        {
            DeleteRenderFolder(folder);
        }

    }
}

internal interface IPowerPointProcessDetector
{
    bool IsPowerPointRunning();
}

internal interface IPowerPointApplicationFactory
{
    bool IsPowerPointInstalled();

    object CreateApplication();
}

internal interface IPowerPointProcessLauncher
{
    void StartPowerPoint();
}

internal sealed record PowerPointWarmStartOptions(
    TimeSpan ProcessStartTimeout,
    TimeSpan ProcessPollInterval,
    TimeSpan AutomationReadyDelay,
    TimeSpan ActivationRetryDelay,
    int MaxActivationAttempts,
    TimeSpan ProcessExitTimeout)
{
    public static PowerPointWarmStartOptions Default { get; } = new(
        TimeSpan.FromSeconds(15),
        TimeSpan.FromMilliseconds(250),
        TimeSpan.FromSeconds(2),
        TimeSpan.FromSeconds(1),
        2,
        TimeSpan.FromSeconds(10));

    public void Validate()
    {
        if (ProcessStartTimeout <= TimeSpan.Zero ||
            ProcessPollInterval <= TimeSpan.Zero ||
            AutomationReadyDelay < TimeSpan.Zero ||
            ActivationRetryDelay < TimeSpan.Zero ||
            MaxActivationAttempts is < 1 or > 3 ||
            ProcessExitTimeout <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(
                nameof(PowerPointWarmStartOptions),
                "PowerPoint warm-start options are outside their supported range.");
        }
    }
}

internal sealed class ComPowerPointApplicationFactory : IPowerPointApplicationFactory
{
    public bool IsPowerPointInstalled() =>
        Type.GetTypeFromProgID("PowerPoint.Application") is not null;

    public object CreateApplication()
    {
        var applicationType = Type.GetTypeFromProgID("PowerPoint.Application")
            ?? throw new InvalidOperationException("Microsoft PowerPoint is not installed.");
        return Activator.CreateInstance(applicationType)
            ?? throw new InvalidOperationException("PowerPoint could not be started.");
    }
}

internal sealed class RegisteredPowerPointProcessLauncher : IPowerPointProcessLauncher
{
    private const string PowerPointAppPath =
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE";

    public void StartPowerPoint()
    {
        var executablePath = FindRegisteredPowerPointExecutable()
            ?? throw new InvalidOperationException(
                "The registered Microsoft PowerPoint executable could not be found.");
        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            UseShellExecute = false,
            WindowStyle = ProcessWindowStyle.Minimized
        };
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Microsoft PowerPoint did not start.");
    }

    private static string? FindRegisteredPowerPointExecutable()
    {
        foreach (var registryView in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView);
            using var appPathKey = baseKey.OpenSubKey(PowerPointAppPath, writable: false);
            if (appPathKey?.GetValue(null) is not string registeredPath)
            {
                continue;
            }

            var expandedPath = Environment.ExpandEnvironmentVariables(registeredPath.Trim().Trim('"'));
            if (!Path.IsPathFullyQualified(expandedPath) ||
                !Path.GetFileName(expandedPath).Equals(
                    "POWERPNT.EXE",
                    StringComparison.OrdinalIgnoreCase) ||
                !File.Exists(expandedPath))
            {
                continue;
            }

            return Path.GetFullPath(expandedPath);
        }

        return null;
    }
}

internal sealed class SystemPowerPointProcessDetector : IPowerPointProcessDetector
{
    public bool IsPowerPointRunning()
    {
        var processes = Process.GetProcessesByName("POWERPNT");
        try
        {
            return processes.Length > 0;
        }
        finally
        {
            foreach (var process in processes)
            {
                process.Dispose();
            }
        }
    }
}
