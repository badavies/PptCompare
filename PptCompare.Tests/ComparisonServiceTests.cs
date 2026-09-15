using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class ComparisonServiceTests
{
    private readonly TextPresentationComparisonService _service = new();

    [TestMethod]
    public async Task ReorderedUnchangedSlidesAreReportedAsMoved()
    {
        var left = CreatePresentation(
            CreateSlide(1, "Current structure", "No content changes"),
            CreateSlide(2, "Final structure", "No content changes"));
        var right = CreatePresentation(
            CreateSlide(1, "Final structure", "No content changes"),
            CreateSlide(2, "Current structure", "No content changes"));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        Assert.AreEqual(0, result.ChangedSlides);
        Assert.AreEqual(2, result.MovedSlides);
        Assert.AreEqual(2, result.Slides.Count(slide => slide.ChangeKind == "Moved"));
    }

    [TestMethod]
    public async Task TextAndImageChangesAreReportedWithoutChangingInput()
    {
        var leftImage = CreateElement("image", SlideElementKind.Image, "left-hash");
        var rightImage = CreateElement("image", SlideElementKind.Image, "right-hash");
        var left = CreatePresentation(CreateSlide(1, "Summary", "Forecast is £42m", leftImage));
        var right = CreatePresentation(CreateSlide(1, "Summary", "Forecast is £44m", rightImage));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        var slide = result.Slides.Single();
        Assert.AreEqual(1, result.ChangedSlides);
        Assert.IsTrue(slide.RightBodySegments?.Any(segment => segment.Kind == DiffKind.Added));
        Assert.IsTrue(slide.ElementChanges.Any(change =>
            change is
            {
                ElementKind: SlideElementKind.Image,
                Kind: SlideElementChangeKind.Replaced
            }));
        Assert.AreEqual("Forecast is £42m", left.Slides[0].Paragraphs[1]);
    }

    [TestMethod]
    public async Task SpatiallyRepeatedTitleChangeCountsAsOneTextBlock()
    {
        var left = CreatePresentation(CreateSpatialSlide(1, "Old title", "Unchanged body"));
        var right = CreatePresentation(CreateSpatialSlide(1, "New title", "Unchanged body"));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        var slide = result.Slides.Single();
        Assert.AreEqual("1 text block changed.", slide.ChangeSummary);
        StringAssert.Contains(slide.LeftBody, "Old title");
        StringAssert.Contains(slide.RightBody, "New title");
        var leftVisibleTitle = Assert.IsInstanceOfType<SlideTextParagraphBlock>(
            slide.LeftBodyContent[0]);
        var rightVisibleTitle = Assert.IsInstanceOfType<SlideTextParagraphBlock>(
            slide.RightBodyContent[0]);
        Assert.AreEqual("Old title", leftVisibleTitle.Paragraph.Text);
        Assert.AreEqual("New title", rightVisibleTitle.Paragraph.Text);
    }

    [TestMethod]
    public async Task UnmarkedBodyMatchingTitleStillCountsAsSeparateTextBlock()
    {
        var left = CreatePresentation(new PresentationSlide(
            1,
            "Old title",
            ["Old title", "Old title"],
            12_192_000,
            6_858_000,
            []));
        var right = CreatePresentation(new PresentationSlide(
            1,
            "New title",
            ["New title", "New title"],
            12_192_000,
            6_858_000,
            []));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        Assert.AreEqual("2 text blocks changed.", result.Slides.Single().ChangeSummary);
    }

    [TestMethod]
    public async Task RepeatedTitleCannotHideEqualValuedBodyAddition()
    {
        var left = CreatePresentation(CreateMarkedSlide(
            "Old title",
            repeatedTitleParagraphIndex: 0,
            "Old title"));
        var right = CreatePresentation(CreateMarkedSlide(
            "New title",
            repeatedTitleParagraphIndex: 1,
            "Old title",
            "New title"));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        var slide = result.Slides.Single();
        Assert.AreEqual("2 text blocks changed.", slide.ChangeSummary);
        Assert.AreEqual(
            "Old title" + Environment.NewLine + Environment.NewLine + "New title",
            string.Concat(slide.RightBodySegments!.Select(segment => segment.Text)));
    }

    [TestMethod]
    public async Task RepeatedTitleCannotTurnEqualValuedBodyIntoFalseChange()
    {
        var left = CreatePresentation(CreateMarkedSlide(
            "Old title",
            repeatedTitleParagraphIndex: 0,
            "Old title",
            "Old title"));
        var right = CreatePresentation(CreateMarkedSlide(
            "New title",
            repeatedTitleParagraphIndex: 1,
            "Old title",
            "New title"));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        var slide = result.Slides.Single();
        Assert.AreEqual("1 text block changed.", slide.ChangeSummary);
        Assert.AreEqual(
            "Old title" + Environment.NewLine + Environment.NewLine + "Old title",
            string.Concat(slide.LeftBodySegments!.Select(segment => segment.Text)));
    }

    [TestMethod]
    public async Task ReorderedLargeCandidateGroupUsesTheTrueClosestElements()
    {
        const int elementCount = 300;
        var leftElements = Enumerable.Range(0, elementCount)
            .Select(index => CreateRepeatedElement($"left-{index}", index))
            .ToArray();
        var rightElements = Enumerable.Range(0, elementCount)
            .Reverse()
            .Select(index => CreateRepeatedElement($"right-{index}", index))
            .ToArray();
        var left = CreatePresentation(CreateSlide(1, "Summary", "Body", leftElements));
        var right = CreatePresentation(CreateSlide(1, "Summary", "Body", rightElements));

        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);

        var slide = result.Slides.Single();
        Assert.AreEqual("Unchanged", slide.ChangeKind);
        Assert.IsEmpty(slide.ElementChanges);
    }

    [TestMethod]
    public async Task CancellationDuringElementCandidateSearchIsObserved()
    {
        const int elementCount = 300;
        var leftElements = Enumerable.Range(0, elementCount)
            .Select(index => CreateRepeatedElement($"left-{index}", index))
            .ToArray();
        var rightBacking = Enumerable.Range(0, elementCount)
            .Reverse()
            .Select(index => CreateRepeatedElement($"right-{index}", index))
            .ToArray();
        using var cancellation = new CancellationTokenSource();
        var rightElements = new CancelOnIndexerReadList<SlideElement>(
            rightBacking,
            elementCount + 10,
            cancellation.Cancel);
        var left = CreatePresentation(CreateSlide(1, "Summary", "Body", leftElements));
        var rightSlide = new PresentationSlide(
            1,
            "Summary",
            ["Summary", "Body"],
            12_192_000,
            6_858_000,
            rightElements);
        var right = CreatePresentation(rightSlide);

        try
        {
            await _service.CompareAsync(left, right, cancellation.Token);
        }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
        {
            Assert.IsGreaterThanOrEqualTo(elementCount + 10, rightElements.IndexerReads);
            return;
        }

        Assert.Fail("Cancellation raised during element matching should stop the comparison.");
    }

    private static LoadedPresentation CreatePresentation(params PresentationSlide[] slides) =>
        new(new PresentationReference("test.pptx", "Test"), slides);

    private static PresentationSlide CreateSlide(
        int number,
        string title,
        string body,
        params SlideElement[] elements) =>
        new(
            number,
            title,
            [title, body],
            12_192_000,
            6_858_000,
            elements);

    private static PresentationSlide CreateSpatialSlide(
        int number,
        string title,
        string body)
    {
        var titleParagraph = new SlideTextParagraph(
            [new SlideTextRun(title, new SlideTextStyle())]);
        var bodyParagraph = new SlideTextParagraph(
            [new SlideTextRun(body, new SlideTextStyle())]);
        return new PresentationSlide(
            number,
            title,
            [title, title, body],
            12_192_000,
            6_858_000,
            [])
        {
            TitleContent = titleParagraph,
            TextContent =
            [
                new SlideTextParagraphBlock(titleParagraph),
                new SlideTextParagraphBlock(bodyParagraph)
            ],
            RepeatedTitleParagraphIndex = 0
        };
    }

    private static PresentationSlide CreateMarkedSlide(
        string title,
        int repeatedTitleParagraphIndex,
        params string[] visibleParagraphs) =>
        new(
            1,
            title,
            [title, .. visibleParagraphs],
            12_192_000,
            6_858_000,
            [])
        {
            RepeatedTitleParagraphIndex = repeatedTitleParagraphIndex
        };

    private static SlideElement CreateElement(
        string id,
        SlideElementKind kind,
        string contentHash) =>
        new(
            id,
            "Test image",
            kind,
            new SlideBounds(100, 100, 500, 500),
            contentHash,
            contentHash,
            string.Empty);

    private static SlideElement CreateRepeatedElement(string id, int position) =>
        new(
            id,
            "Repeated image",
            SlideElementKind.Image,
            new SlideBounds(position * 10_000L, 100, 500, 500),
            "shared-content",
            "shared-visual",
            string.Empty);

    private sealed class CancelOnIndexerReadList<T>(
        IReadOnlyList<T> items,
        int cancelOnRead,
        Action cancel) : IReadOnlyList<T>
    {
        private int _indexerReads;

        public T this[int index]
        {
            get
            {
                var readCount = Interlocked.Increment(ref _indexerReads);
                if (readCount == cancelOnRead)
                {
                    cancel();
                }

                return items[index];
            }
        }

        public int Count => items.Count;

        public int IndexerReads => Volatile.Read(ref _indexerReads);

        public IEnumerator<T> GetEnumerator() => items.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
            GetEnumerator();
    }

    public TestContext TestContext { get; set; } = null!;
}
