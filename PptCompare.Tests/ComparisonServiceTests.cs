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

        var result = await _service.CompareAsync(left, right);

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

        var result = await _service.CompareAsync(left, right);

        var slide = result.Slides.Single();
        Assert.AreEqual(1, result.ChangedSlides);
        Assert.IsTrue(slide.RightBodySegments?.Any(segment => segment.Kind == DiffKind.Added));
        Assert.IsTrue(slide.ElementChanges.Any(change =>
            change.ElementKind == SlideElementKind.Image &&
            change.Kind == SlideElementChangeKind.Replaced));
        Assert.AreEqual("Forecast is £42m", left.Slides[0].Paragraphs[1]);
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
}
