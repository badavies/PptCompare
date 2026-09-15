using System.Text;
using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class WordDiffPerformanceTests
{
    private readonly TextPresentationComparisonService _service = new();

    [TestMethod]
    public async Task LargeSingleKindRunsUseMultipleDirectionChunksAndReconstructExactly()
    {
        const int tokenCount = 600;
        var leftTitle = CreateTokenText(
            tokenCount,
            new string('L', 150),
            new string('L', 150),
            " ");
        var rightTitle = CreateTokenText(
            tokenCount,
            new string('R', 150),
            new string('R', 150),
            "\t");
        var packedByteCount = checked(
            ((long)tokenCount * tokenCount *
             TextPresentationComparisonService.DiffDirectionBitsPerCell + 7) / 8);

        Assert.IsTrue(
            TextPresentationComparisonService.DiffDirectionChunkSizeBytes < 85_000,
            "Each direction chunk must remain below the large-object-heap threshold.");
        Assert.IsTrue(
            packedByteCount > TextPresentationComparisonService.DiffDirectionChunkSizeBytes,
            "This input must exercise more than one direction chunk.");

        var slide = await CompareTitlesAsync(leftTitle, rightTitle);

        AssertSegments(
            slide.LeftTitleSegments,
            leftTitle,
            (leftTitle, DiffKind.Removed));
        AssertSegments(
            slide.RightTitleSegments,
            rightTitle,
            (rightTitle, DiffKind.Added));
    }

    [TestMethod]
    public async Task TokenLimitRetainsGranularDiffAndNextTokenFallsBack()
    {
        const string commonToken = "common";
        var atLimit = CreateTokenText(
            TextPresentationComparisonService.MaxDiffTokensPerSide,
            commonToken,
            "x",
            " ");
        var overLimit = CreateTokenText(
            TextPresentationComparisonService.MaxDiffTokensPerSide + 1,
            commonToken,
            "x",
            " ");

        var atLimitSlide = await CompareTitlesAsync(atLimit, commonToken);
        var overLimitSlide = await CompareTitlesAsync(overLimit, commonToken);

        AssertSegments(
            atLimitSlide.LeftTitleSegments,
            atLimit,
            (commonToken, DiffKind.Unchanged),
            (atLimit[commonToken.Length..], DiffKind.Removed));
        AssertSegments(
            atLimitSlide.RightTitleSegments,
            commonToken,
            (commonToken, DiffKind.Unchanged));
        AssertSegments(
            overLimitSlide.LeftTitleSegments,
            overLimit,
            (overLimit, DiffKind.Removed));
        AssertSegments(
            overLimitSlide.RightTitleSegments,
            commonToken,
            (commonToken, DiffKind.Added));
    }

    [TestMethod]
    public async Task AmbiguousWordDiffPreservesRemovalFirstTieBreaking()
    {
        const string leftTitle = "A B";
        const string rightTitle = "B A";

        var slide = await CompareTitlesAsync(leftTitle, rightTitle);

        AssertSegments(
            slide.LeftTitleSegments,
            leftTitle,
            ("A ", DiffKind.Removed),
            ("B", DiffKind.Unchanged));
        AssertSegments(
            slide.RightTitleSegments,
            rightTitle,
            ("B", DiffKind.Unchanged),
            (" A", DiffKind.Added));
    }

    private async Task<SlideComparisonItem> CompareTitlesAsync(
        string leftTitle,
        string rightTitle)
    {
        var left = CreatePresentation(CreateSlide(leftTitle));
        var right = CreatePresentation(CreateSlide(rightTitle));
        var result = await _service.CompareAsync(left, right, TestContext.CancellationToken);
        return result.Slides.Single();
    }

    private static void AssertSegments(
        IReadOnlyList<DiffSegment>? actual,
        string reconstructedText,
        params (string Text, DiffKind Kind)[] expected)
    {
        Assert.IsNotNull(actual);
        Assert.AreEqual(expected.Length, actual.Count);
        for (var index = 0; index < expected.Length; index++)
        {
            Assert.AreEqual(expected[index].Text, actual[index].Text);
            Assert.AreEqual(expected[index].Kind, actual[index].Kind);
        }

        Assert.AreEqual(reconstructedText, string.Concat(actual.Select(segment => segment.Text)));
    }

    private static string CreateTokenText(
        int tokenCount,
        string firstWord,
        string remainingWord,
        string whitespace)
    {
        var result = new StringBuilder();
        for (var index = 0; index < tokenCount; index++)
        {
            result.Append(index % 2 == 0
                ? index == 0 ? firstWord : remainingWord
                : whitespace);
        }

        return result.ToString();
    }

    private static LoadedPresentation CreatePresentation(PresentationSlide slide) =>
        new(new PresentationReference("test.pptx", "Test"), [slide]);

    private static PresentationSlide CreateSlide(string title) =>
        new(
            1,
            title,
            [title, "Body"],
            12_192_000,
            6_858_000,
            []);

    public TestContext TestContext { get; set; } = null!;
}
