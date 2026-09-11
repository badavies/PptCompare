using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Security;
using PptCompare.Models;

namespace PptCompare.Services;

public sealed class PowerPointPresentationRenderer : IPresentationRenderer
{
    private const int ComServerExecutionFailed = unchecked((int)0x80080005);
    private const int MaxCachedPresentations = 8;
    private const long MaxRenderedImageBytes = 50L * 1024 * 1024;
    private const long MaxRenderedPresentationBytes = 500L * 1024 * 1024;
    private static readonly TimeSpan RenderTimeout = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan StaleRenderFolderAge = TimeSpan.FromHours(24);
    private readonly object _renderFolderGate = new();
    private readonly SemaphoreSlim _renderSemaphore = new(1, 1);
    private readonly Queue<string> _renderFolders = new();
    private readonly IApplicationDiagnostics _diagnostics;
    private readonly string _renderRoot;
    private readonly TimeProvider _timeProvider;
    private bool _disposed;

    public PowerPointPresentationRenderer()
        : this(NullApplicationDiagnostics.Instance)
    {
    }

    public PowerPointPresentationRenderer(IApplicationDiagnostics diagnostics)
        : this(
            diagnostics,
            Path.Combine(Path.GetTempPath(), "PptCompare", "renders"),
            TimeProvider.System)
    {
    }

    internal PowerPointPresentationRenderer(
        IApplicationDiagnostics diagnostics,
        string renderRoot,
        TimeProvider timeProvider)
    {
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        ArgumentException.ThrowIfNullOrWhiteSpace(renderRoot);
        _renderRoot = Path.GetFullPath(renderRoot);
        _timeProvider = timeProvider ?? throw new ArgumentNullException(nameof(timeProvider));
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
            var applicationType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (applicationType is null)
            {
                return new SlideRenderingResult(
                    new Dictionary<int, string>(),
                    "Microsoft PowerPoint was not found; using the built-in slide preview.");
            }

            using var renderCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken);
            var completion = new TaskCompletionSource<SlideRenderingResult>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
                RenderOnStaThread(
                    applicationType,
                    presentationPath,
                    slideCount,
                    completion,
                    renderCancellation.Token))
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
                renderCancellation.Cancel();
                Debug.WriteLine(
                    $"PowerPoint rendering timed out after {RenderTimeout.TotalSeconds:N0} seconds.");
                _diagnostics.RecordEvent("PreviewRenderTimedOut");
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
        Type applicationType,
        string presentationPath,
        int expectedSlideCount,
        TaskCompletionSource<SlideRenderingResult> completion,
        CancellationToken cancellationToken)
    {
        object? application = null;
        object? presentations = null;
        object? presentation = null;
        object? slides = null;
        object? pageSetup = null;
        string? outputFolder = null;

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            outputFolder = CreateRenderFolder();
            application = Activator.CreateInstance(applicationType)
                ?? throw new InvalidOperationException("PowerPoint could not be started.");

            dynamic powerPoint = application;
            powerPoint.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            presentations = powerPoint.Presentations;
            dynamic presentationCollection = presentations;
            presentation = presentationCollection.Open(
                presentationPath,
                -1, // ReadOnly: msoTrue
                0,  // Untitled: msoFalse
                0); // WithWindow: msoFalse

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
            try
            {
                openedPresentation.Export(
                    bulkFolder,
                    "PNG",
                    exportSize.Width,
                    exportSize.Height);
                cancellationToken.ThrowIfCancellationRequested();
                images = CollectExportedImages(bulkFolder, count);
            }
            catch (Exception exception) when (
                exception is COMException or IOException or UnauthorizedAccessException)
            {
                Debug.WriteLine($"Whole-presentation PowerPoint export failed: {exception}");
                _diagnostics.RecordException("PreviewBulkExportFailed", exception);
                images = [];
            }

            if (images.Count != count)
            {
                // Some older or repaired decks fail whole-presentation export even though
                // PowerPoint can render their slides individually.
                DeleteRenderFolder(bulkFolder);
                images = ExportSlidesIndividually(
                    slideCollection,
                    outputFolder,
                    count,
                    exportSize.Width,
                    exportSize.Height,
                    cancellationToken);
            }

            RememberRenderFolder(outputFolder);
            completion.TrySetResult(new SlideRenderingResult(
                images,
                images.Count == count
                    ? "High-fidelity slide previews rendered by Microsoft PowerPoint."
                    : "Some slides could not be rendered; using built-in previews where necessary."));
            outputFolder = null;
        }
        catch (OperationCanceledException)
        {
            completion.TrySetCanceled(cancellationToken);
        }
        catch (Exception exception)
        {
            Debug.WriteLine($"PowerPoint rendering failed: {exception}");
            _diagnostics.RecordException("PreviewRenderFailed", exception);
            completion.TrySetResult(new SlideRenderingResult(
                new Dictionary<int, string>(),
                GetRenderingFailureStatus(exception)));
        }
        finally
        {
            try
            {
                if (presentation is not null)
                {
                    ((dynamic)presentation).Close();
                }
            }
            catch (COMException)
            {
            }

            ReleaseComObject(slides);
            ReleaseComObject(pageSetup);
            ReleaseComObject(presentation);
            ReleaseComObject(presentations);

            try
            {
                if (application is not null)
                {
                    ((dynamic)application).Quit();
                }
            }
            catch (COMException)
            {
            }

            ReleaseComObject(application);
            if (outputFolder is not null)
            {
                DeleteRenderFolder(outputFolder);
            }
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
            foldersToDelete = _renderFolders.ToList();
            _renderFolders.Clear();
        }

        foreach (var folder in foldersToDelete)
        {
            DeleteRenderFolder(folder);
        }

        GC.SuppressFinalize(this);
    }
}
