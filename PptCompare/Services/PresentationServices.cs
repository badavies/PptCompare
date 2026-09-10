using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using PptCompare.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptCompare.Services;

public sealed record PresentationReadLimits(
    long MaxFileBytes = 250L * 1024 * 1024,
    long MaxCharactersPerPart = 5_000_000,
    int MaxSlides = 2_000,
    int MaxParagraphsPerSlide = 10_000,
    int MaxCharactersPerSlide = 500_000,
    long MaxTotalCharacters = 20_000_000);

public sealed class PresentationLoadException : Exception
{
    public PresentationLoadException(string message)
        : base(message)
    {
    }

    public PresentationLoadException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

public interface IFilePickerService
{
    string? PickPresentation();
}

public sealed class WpfFilePickerService : IFilePickerService
{
    public string? PickPresentation()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Open PowerPoint presentation",
            Filter = "PowerPoint presentations (*.pptx;*.pptm)|*.pptx;*.pptm",
            DefaultExt = ".pptx",
            AddExtension = true,
            CheckFileExists = true,
            Multiselect = false,
            ValidateNames = true,
            DereferenceLinks = true
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
}

public interface IPresentationSourceService
{
    Task<LoadedPresentation> LoadAsync(
        PresentationReference presentation,
        CancellationToken cancellationToken = default);
}

public interface IPresentationComparisonService
{
    Task<PresentationComparisonResult> CompareAsync(
        LoadedPresentation left,
        LoadedPresentation right,
        CancellationToken cancellationToken = default);
}

public sealed class OpenXmlPresentationSourceService : IPresentationSourceService
{
    private readonly PresentationReadLimits _limits;

    public OpenXmlPresentationSourceService(PresentationReadLimits? limits = null)
    {
        _limits = limits ?? new PresentationReadLimits();
        ValidateLimits(_limits);
    }

    public Task<LoadedPresentation> LoadAsync(
        PresentationReference presentation,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(presentation);
        return Task.Run(() => Load(presentation, cancellationToken), cancellationToken);
    }

    private LoadedPresentation Load(
        PresentationReference presentation,
        CancellationToken cancellationToken)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = ValidateFile(presentation.Location);
            var openSettings = new OpenSettings
            {
                AutoSave = false,
                MaxCharactersInPart = _limits.MaxCharactersPerPart
            };

            using var document = PresentationDocument.Open(path, false, openSettings);
            var presentationPart = document.PresentationPart
                ?? throw new PresentationLoadException(
                    "The selected file does not contain a readable PowerPoint presentation.");
            var presentationRoot = presentationPart.Presentation
                ?? throw new PresentationLoadException(
                    "The selected file does not contain a readable presentation root.");
            var slideIdList = presentationRoot.SlideIdList
                ?? throw new PresentationLoadException("The selected presentation does not contain any slides.");
            var slideIds = slideIdList.Elements<P.SlideId>().Take(_limits.MaxSlides + 1).ToList();

            if (slideIds.Count > _limits.MaxSlides)
            {
                throw new PresentationLoadException(
                    $"The presentation contains more than the supported limit of {_limits.MaxSlides:N0} slides.");
            }

            var slides = new List<PresentationSlide>();
            long totalCharacters = 0;
            foreach (var slideId in slideIds)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var relationshipId = slideId.RelationshipId?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId) ||
                    presentationPart.GetPartById(relationshipId) is not SlidePart slidePart)
                {
                    continue;
                }

                var slideRoot = slidePart.Slide;
                if (slideRoot is null)
                {
                    continue;
                }

                var paragraphs = new List<string>();
                var paragraphCount = 0;
                var slideCharacters = 0;
                foreach (var paragraph in slideRoot.Descendants<A.Paragraph>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    paragraphCount++;
                    if (paragraphCount > _limits.MaxParagraphsPerSlide)
                    {
                        throw new PresentationLoadException(
                            $"Slide {slides.Count + 1} contains too many text paragraphs to process safely.");
                    }

                    var text = string.Concat(
                        paragraph.Descendants<A.Text>().Select(value => value.Text)).Trim();
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    slideCharacters += text.Length;
                    totalCharacters += text.Length;
                    if (slideCharacters > _limits.MaxCharactersPerSlide ||
                        totalCharacters > _limits.MaxTotalCharacters)
                    {
                        throw new PresentationLoadException(
                            "The presentation contains more text than can be processed safely.");
                    }

                    paragraphs.Add(text);
                }

                if (paragraphs.Count == 0)
                {
                    paragraphs.Add("(No text content on this slide)");
                }

                slides.Add(new PresentationSlide(
                    slides.Count + 1,
                    paragraphs[0],
                    paragraphs));
            }

            if (slides.Count == 0)
            {
                throw new PresentationLoadException("No readable slides were found in the presentation.");
            }

            return new LoadedPresentation(presentation, slides);
        }
        catch (PresentationLoadException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or OpenXmlPackageException or
            XmlException or FormatException or InvalidOperationException or ArgumentException or
            NotSupportedException or System.Security.SecurityException)
        {
            throw new PresentationLoadException(
                "The selected file is damaged, unsupported, or cannot be accessed.",
                exception);
        }
    }

    private string ValidateFile(string location)
    {
        if (string.IsNullOrWhiteSpace(location))
        {
            throw new PresentationLoadException("No presentation file was selected.");
        }

        var extension = Path.GetExtension(location);
        if (!extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".pptm", StringComparison.OrdinalIgnoreCase))
        {
            throw new PresentationLoadException("Select a .pptx or .pptm PowerPoint presentation.");
        }

        var fullPath = Path.GetFullPath(location);
        var file = new FileInfo(fullPath);
        if (!file.Exists)
        {
            throw new PresentationLoadException("The selected presentation no longer exists.");
        }

        if (file.Length == 0)
        {
            throw new PresentationLoadException("The selected presentation is empty.");
        }

        if (file.Length > _limits.MaxFileBytes)
        {
            throw new PresentationLoadException(
                $"The selected presentation is larger than the supported limit of {_limits.MaxFileBytes / 1024 / 1024:N0} MB.");
        }

        return fullPath;
    }

    private static void ValidateLimits(PresentationReadLimits limits)
    {
        if (limits.MaxFileBytes <= 0 ||
            limits.MaxCharactersPerPart <= 0 ||
            limits.MaxSlides <= 0 ||
            limits.MaxParagraphsPerSlide <= 0 ||
            limits.MaxCharactersPerSlide <= 0 ||
            limits.MaxTotalCharacters <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(limits), "All presentation read limits must be positive.");
        }
    }
}

public sealed partial class TextPresentationComparisonService : IPresentationComparisonService
{
    private const int MaxDiffTokensPerSide = 4_000;

    public Task<PresentationComparisonResult> CompareAsync(
        LoadedPresentation left,
        LoadedPresentation right,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        return Task.Run(() =>
        {
            var results = new List<SlideComparisonItem>();
            var total = Math.Max(left.Slides.Count, right.Slides.Count);

            for (var index = 0; index < total; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var leftSlide = index < left.Slides.Count ? left.Slides[index] : null;
                var rightSlide = index < right.Slides.Count ? right.Slides[index] : null;
                results.Add(CreateComparison(index + 1, leftSlide, rightSlide));
            }

            return new PresentationComparisonResult(
                results,
                results.Count(item => item.ChangeKind == "Changed"),
                results.Count(item => item.ChangeKind == "Added"),
                results.Count(item => item.ChangeKind == "Removed"));
        }, cancellationToken);
    }

    private static SlideComparisonItem CreateComparison(
        int number,
        PresentationSlide? left,
        PresentationSlide? right)
    {
        if (left is null)
        {
            return new SlideComparisonItem
            {
                Number = number,
                Title = right!.Title,
                ChangeKind = "Added",
                ChangeColor = "#15803D",
                ChangeSummary = "This slide was added in the right-hand presentation.",
                LeftTitle = "Slide not present",
                LeftBody = "This slide does not exist in the selected left presentation.",
                LeftCallout = "Added on the right",
                RightTitle = right.Title,
                RightBody = FormatBody(right),
                RightCallout = "New slide",
                RightTitleSegments = MarkAll(right.Title, DiffKind.Added),
                RightBodySegments = MarkAll(FormatBody(right), DiffKind.Added)
            };
        }

        if (right is null)
        {
            return new SlideComparisonItem
            {
                Number = number,
                Title = left.Title,
                ChangeKind = "Removed",
                ChangeColor = "#C2413B",
                ChangeSummary = "This slide was removed from the right-hand presentation.",
                LeftTitle = left.Title,
                LeftBody = FormatBody(left),
                LeftCallout = "Removed on the right",
                RightTitle = "Slide not present",
                RightBody = "This slide does not exist in the selected right presentation.",
                RightCallout = "Removed slide",
                LeftTitleSegments = MarkAll(left.Title, DiffKind.Removed),
                LeftBodySegments = MarkAll(FormatBody(left), DiffKind.Removed)
            };
        }

        var leftBody = FormatBody(left);
        var rightBody = FormatBody(right);
        var titleDiff = BuildWordDiff(left.Title, right.Title);
        var bodyDiff = BuildParagraphDiff(
            left.Paragraphs.Skip(1).ToList(),
            right.Paragraphs.Skip(1).ToList());
        var titleChanged = !string.Equals(
            Normalize(left.Title),
            Normalize(right.Title),
            StringComparison.Ordinal);
        var changedBlockCount = bodyDiff.ChangedBlocks + (titleChanged ? 1 : 0);
        var unchanged = changedBlockCount == 0;

        return new SlideComparisonItem
        {
            Number = number,
            Title = right.Title,
            ChangeKind = unchanged ? "Unchanged" : "Changed",
            ChangeColor = unchanged ? "#94A3B8" : "#C77B16",
            ChangeSummary = unchanged
                ? "No text differences detected."
                : $"{changedBlockCount} text block{(changedBlockCount == 1 ? "" : "s")} changed.",
            LeftTitle = left.Title,
            LeftBody = leftBody,
            LeftCallout = unchanged
                ? "No text removed"
                : DescribeDifferences(
                    "Removed or changed",
                    titleDiff.Left.Concat(bodyDiff.Left),
                    DiffKind.Removed),
            RightTitle = right.Title,
            RightBody = rightBody,
            RightCallout = unchanged
                ? "No text added"
                : DescribeDifferences(
                    "Added or changed",
                    titleDiff.Right.Concat(bodyDiff.Right),
                    DiffKind.Added),
            LeftTitleSegments = titleDiff.Left,
            RightTitleSegments = titleDiff.Right,
            LeftBodySegments = bodyDiff.Left,
            RightBodySegments = bodyDiff.Right
        };
    }

    private static string FormatBody(PresentationSlide slide) =>
        string.Join(Environment.NewLine + Environment.NewLine, slide.Paragraphs.Skip(1));

    private static string DescribeDifferences(
        string label,
        IEnumerable<DiffSegment> segments,
        DiffKind kind)
    {
        var values = segments
            .Where(segment => segment.Kind == kind)
            .Select(segment => string.Join(
                ' ',
                segment.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)))
            .Where(value => value.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(3)
            .ToList();

        return values.Count == 0 ? $"{label}: none" : $"{label}: {string.Join(" · ", values)}";
    }

    private static string Normalize(string value) =>
        string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToUpperInvariant();

    private static IReadOnlyList<DiffSegment> MarkAll(string text, DiffKind kind) =>
        string.IsNullOrEmpty(text) ? [] : [new DiffSegment(text, kind)];

    private static (
        IReadOnlyList<DiffSegment> Left,
            IReadOnlyList<DiffSegment> Right,
            int ChangedBlocks) BuildParagraphDiff(
            List<string> leftParagraphs,
            List<string> rightParagraphs)
    {
        var matches = FindMatchingParagraphs(leftParagraphs, rightParagraphs);
        var leftSegments = new List<DiffSegment>();
        var rightSegments = new List<DiffSegment>();
        var changedBlocks = 0;
        var leftStart = 0;
        var rightStart = 0;

        foreach (var match in matches.Append((
                     Left: leftParagraphs.Count,
                     Right: rightParagraphs.Count)))
        {
            var leftCount = match.Left - leftStart;
            var rightCount = match.Right - rightStart;
            var pairedCount = Math.Min(leftCount, rightCount);

            for (var offset = 0; offset < pairedCount; offset++)
            {
                var wordDiff = BuildWordDiff(
                    leftParagraphs[leftStart + offset],
                    rightParagraphs[rightStart + offset]);
                AppendParagraph(leftSegments, wordDiff.Left);
                AppendParagraph(rightSegments, wordDiff.Right);
            }

            for (var offset = pairedCount; offset < leftCount; offset++)
            {
                AppendParagraph(
                    leftSegments,
                    MarkAll(leftParagraphs[leftStart + offset], DiffKind.Removed));
            }

            for (var offset = pairedCount; offset < rightCount; offset++)
            {
                AppendParagraph(
                    rightSegments,
                    MarkAll(rightParagraphs[rightStart + offset], DiffKind.Added));
            }

            changedBlocks += Math.Max(leftCount, rightCount);

            if (match.Left < leftParagraphs.Count && match.Right < rightParagraphs.Count)
            {
                AppendParagraph(
                    leftSegments,
                    MarkAll(leftParagraphs[match.Left], DiffKind.Unchanged));
                AppendParagraph(
                    rightSegments,
                    MarkAll(rightParagraphs[match.Right], DiffKind.Unchanged));
            }

            leftStart = match.Left + 1;
            rightStart = match.Right + 1;
        }

        return (leftSegments, rightSegments, changedBlocks);
    }

    private static List<(int Left, int Right)> FindMatchingParagraphs(
        List<string> leftParagraphs,
        List<string> rightParagraphs)
    {
        const long maxMatrixCells = 250_000;
        if ((long)leftParagraphs.Count * rightParagraphs.Count > maxMatrixCells)
        {
            return [];
        }

        var leftValues = leftParagraphs.Select(Normalize).ToList();
        var rightValues = rightParagraphs.Select(Normalize).ToList();
        var lengths = new int[leftValues.Count + 1, rightValues.Count + 1];

        for (var leftIndex = leftValues.Count - 1; leftIndex >= 0; leftIndex--)
        {
            for (var rightIndex = rightValues.Count - 1; rightIndex >= 0; rightIndex--)
            {
                lengths[leftIndex, rightIndex] =
                    string.Equals(leftValues[leftIndex], rightValues[rightIndex], StringComparison.Ordinal)
                        ? lengths[leftIndex + 1, rightIndex + 1] + 1
                        : Math.Max(lengths[leftIndex + 1, rightIndex], lengths[leftIndex, rightIndex + 1]);
            }
        }

        var matches = new List<(int Left, int Right)>();
        var i = 0;
        var j = 0;
        while (i < leftValues.Count && j < rightValues.Count)
        {
            if (string.Equals(leftValues[i], rightValues[j], StringComparison.Ordinal))
            {
                matches.Add((i++, j++));
            }
            else if (lengths[i + 1, j] >= lengths[i, j + 1])
            {
                i++;
            }
            else
            {
                j++;
            }
        }

        return matches;
    }

    private static void AppendParagraph(
        List<DiffSegment> destination,
        IReadOnlyList<DiffSegment> paragraph)
    {
        if (destination.Count > 0)
        {
            AppendSegment(destination, Environment.NewLine + Environment.NewLine, DiffKind.Unchanged);
        }

        foreach (var segment in paragraph)
        {
            AppendSegment(destination, segment.Text, segment.Kind);
        }
    }

    private static (IReadOnlyList<DiffSegment> Left, IReadOnlyList<DiffSegment> Right) BuildWordDiff(
        string leftText,
        string rightText)
    {
        var leftTokens = Tokenize(leftText);
        var rightTokens = Tokenize(rightText);

        if (leftTokens.Length > MaxDiffTokensPerSide ||
            rightTokens.Length > MaxDiffTokensPerSide ||
            (long)leftTokens.Length * rightTokens.Length > 2_000_000)
        {
            return (MarkAll(leftText, DiffKind.Removed), MarkAll(rightText, DiffKind.Added));
        }

        var lengths = new int[leftTokens.Length + 1, rightTokens.Length + 1];
        for (var leftIndex = leftTokens.Length - 1; leftIndex >= 0; leftIndex--)
        {
            for (var rightIndex = rightTokens.Length - 1; rightIndex >= 0; rightIndex--)
            {
                lengths[leftIndex, rightIndex] = TokensEqual(leftTokens[leftIndex], rightTokens[rightIndex])
                    ? lengths[leftIndex + 1, rightIndex + 1] + 1
                    : Math.Max(lengths[leftIndex + 1, rightIndex], lengths[leftIndex, rightIndex + 1]);
            }
        }

        var leftSegments = new List<DiffSegment>();
        var rightSegments = new List<DiffSegment>();
        var i = 0;
        var j = 0;

        while (i < leftTokens.Length || j < rightTokens.Length)
        {
            if (i < leftTokens.Length && j < rightTokens.Length &&
                TokensEqual(leftTokens[i], rightTokens[j]))
            {
                AppendSegment(leftSegments, leftTokens[i++], DiffKind.Unchanged);
                AppendSegment(rightSegments, rightTokens[j++], DiffKind.Unchanged);
            }
            else if (i < leftTokens.Length &&
                     (j == rightTokens.Length || lengths[i + 1, j] >= lengths[i, j + 1]))
            {
                AppendSegment(leftSegments, leftTokens[i++], DiffKind.Removed);
            }
            else
            {
                AppendSegment(rightSegments, rightTokens[j++], DiffKind.Added);
            }
        }

        return (leftSegments, rightSegments);
    }

    private static string[] Tokenize(string text) =>
        DiffTokenRegex().Matches(text)
            .Cast<Match>()
            .Select(match => match.Value)
            .ToArray();

    [GeneratedRegex(@"\s+|[^\s]+", RegexOptions.NonBacktracking)]
    private static partial Regex DiffTokenRegex();

    private static bool TokensEqual(string left, string right) =>
        string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

    private static void AppendSegment(List<DiffSegment> segments, string text, DiffKind kind)
    {
        if (segments.Count > 0 && segments[^1].Kind == kind)
        {
            segments[^1] = segments[^1] with { Text = segments[^1].Text + text };
            return;
        }

        segments.Add(new DiffSegment(text, kind));
    }
}
