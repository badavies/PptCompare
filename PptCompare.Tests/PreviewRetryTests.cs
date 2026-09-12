using System.Collections.Concurrent;
using System.Diagnostics;
using PptCompare.Models;
using PptCompare.Services;
using PptCompare.ViewModels;

namespace PptCompare.Tests;

[TestClass]
public sealed class PreviewRetryTests
{
    private static readonly string[] ExpectedRenderedPaths =
        ["left.pptx", "right.pptx", "left.pptx", "right.pptx"];

    [TestMethod]
    public async Task CompareRetriesOnlyMissingHighFidelityPreviews()
    {
        var renderer = new ControllableRenderer();
        var filePicker = new QueueFilePicker("left.pptx", "right.pptx");
        using var viewModel = new MainWindowViewModel(
            filePicker,
            new StubPresentationSource(),
            new TextPresentationComparisonService(),
            renderer,
            new RecordingDiagnostics(),
            new StubSettingsService(),
            new StubSettingsDialogService(),
            new StubAboutDialogService(),
            new ApplicationSettings(UsePowerPointRendering: true));

        viewModel.OpenPresentationCommand.Execute("Left");
        await WaitUntilAsync(() =>
            viewModel.LeftVersions.Count == 1 && renderer.CallCount >= 1);
        viewModel.OpenPresentationCommand.Execute("Right");
        await WaitUntilAsync(() =>
            viewModel.RightVersions.Count == 1 && renderer.CallCount >= 2);

        renderer.HighFidelityAvailable = true;
        viewModel.CompareCommand.Execute(null);
        await WaitUntilAsync(() => renderer.CallCount >= 4 && viewModel.Slides.Count == 1);

        var callsAfterSuccessfulRetry = renderer.CallCount;
        viewModel.CompareCommand.Execute(null);
        await WaitUntilAsync(() => viewModel.StatusMessage.StartsWith(
            "Comparison complete",
            StringComparison.Ordinal));
        await Task.Delay(50);

        Assert.AreEqual(4, callsAfterSuccessfulRetry);
        Assert.AreEqual(callsAfterSuccessfulRetry, renderer.CallCount);
        CollectionAssert.AreEqual(ExpectedRenderedPaths, renderer.RenderedPaths.ToArray());
    }

    private static async Task WaitUntilAsync(Func<bool> condition)
    {
        var timeout = Stopwatch.StartNew();
        while (!condition() && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(10);
        }

        Assert.IsTrue(condition(), "The asynchronous preview operation did not finish in time.");
    }

    private static LoadedPresentation CreatePresentation(PresentationReference source) =>
        new(
            source,
            [
                new PresentationSlide(
                    1,
                    "Summary",
                    ["Summary", "No changes"],
                    12_192_000,
                    6_858_000,
                    [])
            ]);

    private sealed class QueueFilePicker(params string[] paths) : IFilePickerService
    {
        private readonly Queue<string> _paths = new(paths);

        public string? PickPresentation() => _paths.Count == 0 ? null : _paths.Dequeue();
    }

    private sealed class StubPresentationSource : IPresentationSourceService
    {
        public Task<LoadedPresentation> LoadAsync(
            PresentationReference presentation,
            CancellationToken cancellationToken = default) =>
            Task.FromResult(CreatePresentation(presentation));
    }

    private sealed class ControllableRenderer : IPresentationRenderer
    {
        private int _callCount;

        public bool HighFidelityAvailable { get; set; }

        public int CallCount => Volatile.Read(ref _callCount);

        public ConcurrentQueue<string> RenderedPaths { get; } = new();

        public Task<SlideRenderingResult> RenderAsync(
            string presentationPath,
            int slideCount,
            CancellationToken cancellationToken = default)
        {
            Interlocked.Increment(ref _callCount);
            RenderedPaths.Enqueue(presentationPath);
            IReadOnlyDictionary<int, string> images = HighFidelityAvailable
                ? new Dictionary<int, string> { [1] = $"preview-{CallCount}.png" }
                : new Dictionary<int, string>();
            return Task.FromResult(new SlideRenderingResult(
                images,
                HighFidelityAvailable ? "Rendered." : "PowerPoint is in use."));
        }

        public void Dispose()
        {
        }
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

    private sealed class StubSettingsService : IApplicationSettingsService
    {
        public ApplicationSettings Load() => new();

        public void Save(ApplicationSettings settings)
        {
        }
    }

    private sealed class StubSettingsDialogService : ISettingsDialogService
    {
        public ApplicationSettings? EditSettings(ApplicationSettings currentSettings) => null;
    }

    private sealed class StubAboutDialogService : IAboutDialogService
    {
        public void ShowAbout()
        {
        }
    }
}
