using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using PptCompare.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptCompare.Services;

public sealed record PresentationReadLimits(
    long MaxFileBytes = 250L * 1024 * 1024,
    long MaxCharactersPerPart = 5_000_000,
    int MaxSlides = 2_000,
    int MaxParagraphsPerSlide = 10_000,
    int MaxCharactersPerSlide = 500_000,
    long MaxTotalCharacters = 20_000_000,
    long MaxRelatedPartBytes = 100L * 1024 * 1024,
    int MaxPreviewImageBytes = 25 * 1024 * 1024,
    long MaxTotalPreviewImageBytes = 100L * 1024 * 1024);

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

public sealed record SlideRenderingResult(
    IReadOnlyDictionary<int, string> SlideImages,
    string Status);

public interface IPresentationRenderer : IDisposable
{
    Task<SlideRenderingResult> RenderAsync(
        string presentationPath,
        int slideCount,
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
    private static readonly XNamespace PresentationNamespace =
        "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace =
        "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private readonly PresentationReadLimits _limits;

    public OpenXmlPresentationSourceService(PresentationReadLimits? limits = null)
    {
        _limits = limits ?? new PresentationReadLimits();
        ValidateLimits(_limits);
    }

    public async Task<LoadedPresentation> LoadAsync(
        PresentationReference presentation,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(presentation);
        return await Task.Run(
            () => Load(presentation, cancellationToken),
            cancellationToken);
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
            long totalPreviewImageBytes = 0;
            var slideWidth = presentationRoot.SlideSize?.Cx?.Value ?? 12_192_000L;
            var slideHeight = presentationRoot.SlideSize?.Cy?.Value ?? 6_858_000L;
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

                var slideXml = XElement.Parse(slideRoot.OuterXml, LoadOptions.None);
                var paragraphCount = slideXml
                    .Descendants(DrawingNamespace + "p")
                    .Take(_limits.MaxParagraphsPerSlide + 1)
                    .Count();
                if (paragraphCount > _limits.MaxParagraphsPerSlide)
                {
                    throw new PresentationLoadException(
                        $"Slide {slides.Count + 1} contains too many text paragraphs to process safely.");
                }

                var paragraphs = ExtractParagraphsInReadingOrder(slidePart, slideXml);
                var slideCharacters = 0;
                foreach (var text in paragraphs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    slideCharacters += text.Length;
                    totalCharacters += text.Length;
                    if (slideCharacters > _limits.MaxCharactersPerSlide ||
                        totalCharacters > _limits.MaxTotalCharacters)
                    {
                        throw new PresentationLoadException(
                            "The presentation contains more text than can be processed safely.");
                    }
                }

                if (paragraphs.Count == 0)
                {
                    paragraphs.Add("(No text content on this slide)");
                }

                var elements = ExtractElements(
                    slidePart,
                    slideXml,
                    ref totalPreviewImageBytes,
                    cancellationToken);
                slides.Add(new PresentationSlide(
                    slides.Count + 1,
                    paragraphs[0],
                    paragraphs,
                    slideWidth,
                    slideHeight,
                    elements));
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
            limits.MaxTotalCharacters <= 0 ||
            limits.MaxRelatedPartBytes <= 0 ||
            limits.MaxPreviewImageBytes <= 0 ||
            limits.MaxTotalPreviewImageBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(limits), "All presentation read limits must be positive.");
        }
    }

    private static List<string> ExtractParagraphsInReadingOrder(
        SlidePart slidePart,
        XElement slideXml)
    {
        var shapeTree = slideXml.Descendants(PresentationNamespace + "spTree").FirstOrDefault();
        if (shapeTree is null)
        {
            return ExtractParagraphs(slideXml);
        }

        var layoutXml = slidePart.SlideLayoutPart?.SlideLayout is { } layout
            ? XElement.Parse(layout.OuterXml, LoadOptions.None)
            : null;
        var masterXml = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster is { } master
            ? XElement.Parse(master.OuterXml, LoadOptions.None)
            : null;
        var candidates = shapeTree
            .Elements()
            .Where(IsDrawableElement)
            .Select((element, index) => new TextReadingCandidate(
                index,
                ResolveBounds(element, layoutXml, masterXml),
                IsTitlePlaceholder(element),
                ExtractParagraphs(element)))
            .Where(candidate => candidate.Paragraphs.Count > 0)
            .ToList();
        if (candidates.Count == 0)
        {
            return ExtractParagraphs(slideXml);
        }

        // PowerPoint stores shapes in drawing/z-order, which is not necessarily reading order.
        // The first entry is metadata for the comparison heading. Keep every actual paragraph,
        // including that title text at its visible position, in the spatially ordered body.
        var title = candidates.FirstOrDefault(candidate => candidate.IsTitle) ?? candidates[0];
        var paragraphs = new List<string> { title.Paragraphs[0] };
        paragraphs.AddRange(candidates
            .OrderBy(candidate => candidate.HasBounds ? 0 : 1)
            .ThenBy(candidate => candidate.Bounds.Y)
            .ThenBy(candidate => candidate.Bounds.X)
            .ThenBy(candidate => candidate.DocumentOrder)
            .SelectMany(candidate => candidate.Paragraphs));
        return paragraphs;
    }

    private static SlideBounds ResolveBounds(
        XElement element,
        XElement? layoutXml,
        XElement? masterXml)
    {
        var bounds = ReadBounds(element);
        if (HasBounds(bounds))
        {
            return bounds;
        }

        var placeholder = element
            .Descendants(PresentationNamespace + "ph")
            .FirstOrDefault();
        if (placeholder is null)
        {
            return bounds;
        }

        var placeholderIndex = placeholder.Attribute("idx")?.Value;
        var placeholderType = placeholder.Attribute("type")?.Value;
        foreach (var inheritedRoot in new[] { layoutXml, masterXml })
        {
            if (inheritedRoot is null)
            {
                continue;
            }

            var inheritedElement = inheritedRoot
                .Descendants()
                .Where(IsDrawableElement)
                .FirstOrDefault(candidate => PlaceholderMatches(
                    candidate,
                    placeholderIndex,
                    placeholderType));
            if (inheritedElement is null)
            {
                continue;
            }

            var inheritedBounds = ReadBounds(inheritedElement);
            if (HasBounds(inheritedBounds))
            {
                return inheritedBounds;
            }
        }

        return bounds;
    }

    private static bool PlaceholderMatches(
        XElement candidate,
        string? expectedIndex,
        string? expectedType)
    {
        var placeholder = candidate
            .Descendants(PresentationNamespace + "ph")
            .FirstOrDefault();
        if (placeholder is null)
        {
            return false;
        }

        var candidateIndex = placeholder.Attribute("idx")?.Value;
        if (!string.IsNullOrWhiteSpace(expectedIndex))
        {
            return string.Equals(candidateIndex, expectedIndex, StringComparison.Ordinal);
        }

        var candidateType = placeholder.Attribute("type")?.Value;
        return string.Equals(candidateType, expectedType, StringComparison.Ordinal);
    }

    private static bool HasBounds(SlideBounds bounds) =>
        bounds.Width > 0 || bounds.Height > 0;

    private static List<string> ExtractParagraphs(XElement element) =>
        element
            .Descendants(DrawingNamespace + "p")
            .Select(paragraph => string.Concat(
                paragraph.Descendants(DrawingNamespace + "t").Select(value => value.Value)).Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();

    private static bool IsTitlePlaceholder(XElement element)
    {
        var placeholderType = element
            .Descendants(PresentationNamespace + "ph")
            .FirstOrDefault()?
            .Attribute("type")?
            .Value;
        return placeholderType is "title" or "ctrTitle";
    }

    private List<SlideElement> ExtractElements(
        SlidePart slidePart,
        XElement slideXml,
        ref long totalPreviewImageBytes,
        CancellationToken cancellationToken)
    {
        var shapeTree = slideXml.Descendants(PresentationNamespace + "spTree").FirstOrDefault();
        if (shapeTree is null)
        {
            return [];
        }

        var elements = new List<SlideElement>();
        var layoutXml = slidePart.SlideLayoutPart?.SlideLayout is { } layout
            ? XElement.Parse(layout.OuterXml, LoadOptions.None)
            : null;
        var masterXml = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster is { } master
            ? XElement.Parse(master.OuterXml, LoadOptions.None)
            : null;
        var fallbackId = 0;
        foreach (var element in shapeTree.Elements().Where(IsDrawableElement))
        {
            cancellationToken.ThrowIfCancellationRequested();
            fallbackId++;
            var kind = GetElementKind(element);
            var drawingProperties = element
                .Descendants(PresentationNamespace + "cNvPr")
                .FirstOrDefault();
            var id = drawingProperties?.Attribute("id")?.Value ?? $"generated-{fallbackId}";
            var name = drawingProperties?.Attribute("name")?.Value ?? $"{kind} {fallbackId}";
            var bounds = ResolveBounds(element, layoutXml, masterXml);
            var text = string.Join(
                Environment.NewLine,
                element.Descendants(DrawingNamespace + "t").Select(value => value.Value));
            var relatedContent = ReadRelatedContent(
                slidePart,
                element,
                kind == SlideElementKind.Image,
                ref totalPreviewImageBytes,
                cancellationToken);
            var canonicalMarkup = CreateCanonicalMarkup(element);
            var contentHash = relatedContent.Hash.Length > 0
                ? relatedContent.Hash
                : HashText(string.Empty);
            var visualHash = HashText(canonicalMarkup + "|" + contentHash);

            elements.Add(new SlideElement(
                id,
                name,
                kind,
                bounds,
                contentHash,
                visualHash,
                text,
                relatedContent.PreviewBytes));
        }

        return elements;
    }

    private static bool IsDrawableElement(XElement element) =>
        element.Name.Namespace == PresentationNamespace &&
        element.Name.LocalName is "sp" or "pic" or "graphicFrame" or "cxnSp" or "grpSp";

    private static SlideElementKind GetElementKind(XElement element) =>
        element.Name.LocalName switch
        {
            "pic" => SlideElementKind.Image,
            "cxnSp" => SlideElementKind.Connector,
            "grpSp" => SlideElementKind.Group,
            "graphicFrame" when element.Descendants(DrawingNamespace + "tbl").Any() =>
                SlideElementKind.Table,
            "graphicFrame" when element.Descendants().Any(value => value.Name.LocalName == "chart") =>
                SlideElementKind.Chart,
            "graphicFrame" => SlideElementKind.Other,
            "sp" when element.Descendants(DrawingNamespace + "t").Any() =>
                SlideElementKind.Text,
            "sp" => SlideElementKind.Shape,
            _ => SlideElementKind.Other
        };

    private static SlideBounds ReadBounds(XElement element)
    {
        var transform = element
            .Descendants()
            .FirstOrDefault(value => value.Name.LocalName == "xfrm");
        var offset = transform?.Elements().FirstOrDefault(value => value.Name.LocalName == "off");
        var extents = transform?.Elements().FirstOrDefault(value => value.Name.LocalName == "ext");

        return new SlideBounds(
            ReadLong(offset, "x"),
            ReadLong(offset, "y"),
            ReadLong(extents, "cx"),
            ReadLong(extents, "cy"));
    }

    private static long ReadLong(XElement? element, string attributeName) =>
        long.TryParse(
            element?.Attribute(attributeName)?.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out var value)
            ? value
            : 0;

    private RelatedContent ReadRelatedContent(
        SlidePart slidePart,
        XElement element,
        bool capturePreview,
        ref long totalPreviewImageBytes,
        CancellationToken cancellationToken)
    {
        var relationshipIds = element
            .DescendantsAndSelf()
            .Attributes()
            .Where(attribute => attribute.Name.Namespace == RelationshipNamespace)
            .Select(attribute => attribute.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .Order(StringComparer.Ordinal)
            .ToList();
        var hashes = new List<string>();
        byte[]? previewBytes = null;

        foreach (var relationshipId in relationshipIds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            OpenXmlPart part;
            try
            {
                part = slidePart.GetPartById(relationshipId);
            }
            catch (ArgumentOutOfRangeException)
            {
                continue;
            }

            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            var partContent = ReadAndHashPart(
                stream,
                capturePreview && previewBytes is null,
                ref totalPreviewImageBytes,
                cancellationToken);
            hashes.Add(partContent.Hash);
            previewBytes ??= partContent.PreviewBytes;
        }

        return new RelatedContent(
            hashes.Count == 0 ? string.Empty : HashText(string.Join("|", hashes)),
            previewBytes);
    }

    private RelatedContent ReadAndHashPart(
        Stream stream,
        bool capturePreview,
        ref long totalPreviewImageBytes,
        CancellationToken cancellationToken)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        using var preview = capturePreview ? new MemoryStream() : null;
        var buffer = new byte[81_920];
        long totalBytes = 0;
        var keepPreview = capturePreview;

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = stream.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                break;
            }

            totalBytes += read;
            if (totalBytes > _limits.MaxRelatedPartBytes)
            {
                throw new PresentationLoadException(
                    "The presentation contains an embedded item that is too large to process safely.");
            }

            hash.AppendData(buffer, 0, read);
            if (keepPreview)
            {
                if (totalBytes <= _limits.MaxPreviewImageBytes &&
                    totalPreviewImageBytes + totalBytes <= _limits.MaxTotalPreviewImageBytes)
                {
                    preview!.Write(buffer, 0, read);
                }
                else
                {
                    keepPreview = false;
                }
            }
        }

        byte[]? previewBytes = null;
        if (keepPreview && preview is not null)
        {
            previewBytes = preview.ToArray();
            totalPreviewImageBytes += previewBytes.Length;
        }

        return new RelatedContent(Convert.ToHexString(hash.GetHashAndReset()), previewBytes);
    }

    private static string CreateCanonicalMarkup(XElement element)
    {
        var clone = new XElement(element);
        foreach (var transform in clone
                     .Descendants()
                     .Where(value => value.Name.LocalName == "xfrm")
                     .ToList())
        {
            transform.Remove();
        }

        foreach (var textValue in clone.Descendants(DrawingNamespace + "t"))
        {
            textValue.Value = string.Empty;
        }

        foreach (var attribute in clone
                     .DescendantsAndSelf()
                     .Attributes()
                     .Where(attribute =>
                         attribute.Name.Namespace == RelationshipNamespace ||
                         attribute.Name.LocalName is "id" or "name")
                     .ToList())
        {
            attribute.Remove();
        }

        return clone.ToString(SaveOptions.DisableFormatting);
    }

    private static string HashText(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));

    private sealed record TextReadingCandidate(
        int DocumentOrder,
        SlideBounds Bounds,
        bool IsTitle,
        List<string> Paragraphs)
    {
        public bool HasBounds => Bounds.Width > 0 || Bounds.Height > 0;
    }

    private sealed record RelatedContent(string Hash, byte[]? PreviewBytes);
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
            var matches = MatchSlides(left.Slides, right.Slides, cancellationToken);
            var matchedLeft = new HashSet<int>();

            for (var rightIndex = 0; rightIndex < right.Slides.Count; rightIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!matches.TryGetValue(rightIndex, out var leftIndex))
                {
                    results.Add(CreateComparison(rightIndex + 1, null, right.Slides[rightIndex]));
                    continue;
                }

                matchedLeft.Add(leftIndex);
                var comparison = CreateComparison(
                    rightIndex + 1,
                    left.Slides[leftIndex],
                    right.Slides[rightIndex]);
                results.Add(leftIndex == rightIndex
                    ? comparison
                    : MarkAsMoved(comparison, leftIndex + 1, rightIndex + 1));
            }

            for (var leftIndex = 0; leftIndex < left.Slides.Count; leftIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!matchedLeft.Contains(leftIndex))
                {
                    results.Add(CreateComparison(results.Count + 1, left.Slides[leftIndex], null));
                }
            }

            return new PresentationComparisonResult(
                results,
                results.Count(item => item.ChangeKind == "Changed"),
                results.Count(item => item.ChangeKind == "Added"),
                results.Count(item => item.ChangeKind == "Removed"),
                results.Count(item => item.ChangeKind == "Moved"));
        }, cancellationToken);
    }

    private static Dictionary<int, int> MatchSlides(
        IReadOnlyList<PresentationSlide> leftSlides,
        IReadOnlyList<PresentationSlide> rightSlides,
        CancellationToken cancellationToken)
    {
        var matches = new Dictionary<int, int>();
        var unmatchedLeft = new HashSet<int>(Enumerable.Range(0, leftSlides.Count));
        var leftFingerprints = leftSlides.Select(CreateSlideFingerprint).ToList();
        var rightFingerprints = rightSlides.Select(CreateSlideFingerprint).ToList();

        // Match identical slides first. This identifies pure reordering without
        // confusing it with edits to slides that happen to occupy the same position.
        for (var rightIndex = 0; rightIndex < rightSlides.Count; rightIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var exactMatches = unmatchedLeft
                .Where(leftIndex => string.Equals(
                    leftFingerprints[leftIndex],
                    rightFingerprints[rightIndex],
                    StringComparison.Ordinal))
                .OrderBy(leftIndex => leftIndex == rightIndex ? 0 : 1)
                .ThenBy(leftIndex => Math.Abs(leftIndex - rightIndex))
                .ToList();
            if (exactMatches.Count == 0)
            {
                continue;
            }

            matches[rightIndex] = exactMatches[0];
            unmatchedLeft.Remove(exactMatches[0]);
        }

        // Pair the remaining edited slides by title and proximity. Any right-hand
        // slide left without a partner is added; unmatched left slides are removed.
        for (var rightIndex = 0; rightIndex < rightSlides.Count; rightIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (matches.ContainsKey(rightIndex) || unmatchedLeft.Count == 0)
            {
                continue;
            }

            var normalizedTitle = Normalize(rightSlides[rightIndex].Title);
            var sameTitle = unmatchedLeft
                .Where(leftIndex => string.Equals(
                    Normalize(leftSlides[leftIndex].Title),
                    normalizedTitle,
                    StringComparison.Ordinal))
                .OrderBy(leftIndex => Math.Abs(leftIndex - rightIndex))
                .FirstOrDefault(-1);
            var leftMatch = sameTitle >= 0
                ? sameTitle
                : unmatchedLeft.Contains(rightIndex)
                    ? rightIndex
                    : unmatchedLeft.OrderBy(leftIndex => Math.Abs(leftIndex - rightIndex)).First();
            matches[rightIndex] = leftMatch;
            unmatchedLeft.Remove(leftMatch);
        }

        return matches;
    }

    private static string CreateSlideFingerprint(PresentationSlide slide)
    {
        var value = new StringBuilder();
        value.Append(Normalize(slide.Title));
        foreach (var paragraph in slide.Paragraphs.Skip(1))
        {
            value.Append('|').Append(Normalize(paragraph));
        }

        foreach (var element in slide.Elements
                     .OrderBy(element => element.Kind)
                     .ThenBy(element => element.Bounds.Y)
                     .ThenBy(element => element.Bounds.X)
                     .ThenBy(element => element.ContentHash, StringComparer.Ordinal))
        {
            value
                .Append('|').Append(element.Kind)
                .Append(':').Append(element.Bounds.X)
                .Append(':').Append(element.Bounds.Y)
                .Append(':').Append(element.Bounds.Width)
                .Append(':').Append(element.Bounds.Height)
                .Append(':').Append(element.ContentHash)
                .Append(':').Append(element.VisualHash);
        }

        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value.ToString())));
    }

    private static SlideComparisonItem MarkAsMoved(
        SlideComparisonItem item,
        int previousPosition,
        int currentPosition)
    {
        var movement = $"Slide moved from position {previousPosition} to {currentPosition}.";
        var movedOnly = item.ChangeKind == "Unchanged";
        return new SlideComparisonItem
        {
            Number = item.Number,
            Title = item.Title,
            ChangeKind = movedOnly ? "Moved" : item.ChangeKind,
            ChangeColor = movedOnly ? "#2563EB" : item.ChangeColor,
            ChangeSummary = movedOnly ? movement : $"{movement} {item.ChangeSummary}",
            LeftTitle = item.LeftTitle,
            LeftBody = item.LeftBody,
            LeftCallout = item.LeftCallout,
            RightTitle = item.RightTitle,
            RightBody = item.RightBody,
            RightCallout = item.RightCallout,
            LeftTitleSegments = item.LeftTitleSegments,
            LeftBodySegments = item.LeftBodySegments,
            RightTitleSegments = item.RightTitleSegments,
            RightBodySegments = item.RightBodySegments,
            LeftSlide = item.LeftSlide,
            RightSlide = item.RightSlide,
            ElementChanges = item.ElementChanges
        };
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
                RightBodySegments = MarkAll(FormatBody(right), DiffKind.Added),
                RightSlide = right,
                ElementChanges = right.Elements
                    .Select(element => new SlideElementChange(
                        SlideElementChangeKind.Added,
                        element.Kind,
                        $"{GetElementLabel(element.Kind)} added",
                        null,
                        element.Bounds))
                    .ToList()
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
                LeftBodySegments = MarkAll(FormatBody(left), DiffKind.Removed),
                LeftSlide = left,
                ElementChanges = left.Elements
                    .Select(element => new SlideElementChange(
                        SlideElementChangeKind.Removed,
                        element.Kind,
                        $"{GetElementLabel(element.Kind)} removed",
                        element.Bounds,
                        null))
                    .ToList()
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
        var elementChanges = CompareElements(left.Elements, right.Elements);
        var unchanged = changedBlockCount == 0 && elementChanges.Count == 0;
        var summary = CreateChangeSummary(changedBlockCount, elementChanges.Count);

        return new SlideComparisonItem
        {
            Number = number,
            Title = right.Title,
            ChangeKind = unchanged ? "Unchanged" : "Changed",
            ChangeColor = unchanged ? "#94A3B8" : "#C77B16",
            ChangeSummary = summary,
            LeftTitle = left.Title,
            LeftBody = leftBody,
            LeftCallout = changedBlockCount > 0
                ? DescribeDifferences(
                    "Removed or changed",
                    titleDiff.Left.Concat(bodyDiff.Left),
                    DiffKind.Removed)
                : DescribeVisualChanges(elementChanges),
            RightTitle = right.Title,
            RightBody = rightBody,
            RightCallout = changedBlockCount > 0
                ? DescribeDifferences(
                    "Added or changed",
                    titleDiff.Right.Concat(bodyDiff.Right),
                    DiffKind.Added)
                : DescribeVisualChanges(elementChanges),
            LeftTitleSegments = titleDiff.Left,
            RightTitleSegments = titleDiff.Right,
            LeftBodySegments = bodyDiff.Left,
            RightBodySegments = bodyDiff.Right,
            LeftSlide = left,
            RightSlide = right,
            ElementChanges = elementChanges
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

    private static string CreateChangeSummary(int changedTextBlocks, int changedElements)
    {
        if (changedTextBlocks == 0 && changedElements == 0)
        {
            return "No text, image, or shape differences detected.";
        }

        var parts = new List<string>();
        if (changedTextBlocks > 0)
        {
            parts.Add(
                $"{changedTextBlocks} text block{(changedTextBlocks == 1 ? "" : "s")}");
        }

        if (changedElements > 0)
        {
            parts.Add(
                $"{changedElements} visual element{(changedElements == 1 ? "" : "s")}");
        }

        return $"{string.Join(" and ", parts)} changed.";
    }

    private static string DescribeVisualChanges(List<SlideElementChange> changes) =>
        changes.Count == 0
            ? "No visual changes detected"
            : string.Join(" · ", changes.Take(3).Select(change => change.Description));

    private static List<SlideElementChange> CompareElements(
        IReadOnlyList<SlideElement> leftElements,
        IReadOnlyList<SlideElement> rightElements)
    {
        var changes = new List<SlideElementChange>();
        var unmatchedRight = new HashSet<int>(Enumerable.Range(0, rightElements.Count));

        foreach (var leftElement in leftElements)
        {
            var rightIndex = FindMatchingElement(leftElement, rightElements, unmatchedRight);
            if (rightIndex < 0)
            {
                changes.Add(new SlideElementChange(
                    SlideElementChangeKind.Removed,
                    leftElement.Kind,
                    $"{GetElementLabel(leftElement.Kind)} removed",
                    leftElement.Bounds,
                    null));
                continue;
            }

            unmatchedRight.Remove(rightIndex);
            var rightElement = rightElements[rightIndex];
            var change = CompareMatchedElement(leftElement, rightElement);
            if (change is not null)
            {
                changes.Add(change);
            }
        }

        foreach (var rightIndex in unmatchedRight.Order())
        {
            var rightElement = rightElements[rightIndex];
            changes.Add(new SlideElementChange(
                SlideElementChangeKind.Added,
                rightElement.Kind,
                $"{GetElementLabel(rightElement.Kind)} added",
                null,
                rightElement.Bounds));
        }

        return changes;
    }

    private static int FindMatchingElement(
        SlideElement left,
        IReadOnlyList<SlideElement> rightElements,
        HashSet<int> candidates)
    {
        var exactId = candidates.FirstOrDefault(
            index => rightElements[index].Kind == left.Kind &&
                     string.Equals(rightElements[index].Id, left.Id, StringComparison.Ordinal),
            -1);
        if (exactId >= 0)
        {
            return exactId;
        }

        var sameContent = candidates
            .Where(index => rightElements[index].Kind == left.Kind &&
                            string.Equals(
                                rightElements[index].ContentHash,
                                left.ContentHash,
                                StringComparison.Ordinal))
            .OrderBy(index => BoundsDistance(left.Bounds, rightElements[index].Bounds))
            .FirstOrDefault(-1);
        if (sameContent >= 0)
        {
            return sameContent;
        }

        var sameName = candidates
            .Where(index => rightElements[index].Kind == left.Kind &&
                            string.Equals(
                                rightElements[index].Name,
                                left.Name,
                                StringComparison.OrdinalIgnoreCase))
            .OrderBy(index => BoundsDistance(left.Bounds, rightElements[index].Bounds))
            .FirstOrDefault(-1);
        if (sameName >= 0)
        {
            return sameName;
        }

        return candidates
            .Where(index => rightElements[index].Kind == left.Kind)
            .OrderBy(index => BoundsDistance(left.Bounds, rightElements[index].Bounds))
            .FirstOrDefault(-1);
    }

    private static SlideElementChange? CompareMatchedElement(
        SlideElement left,
        SlideElement right)
    {
        var moved = left.Bounds.X != right.Bounds.X || left.Bounds.Y != right.Bounds.Y;
        var resized = left.Bounds.Width != right.Bounds.Width ||
                      left.Bounds.Height != right.Bounds.Height;
        var contentChanged = !string.Equals(
            left.ContentHash,
            right.ContentHash,
            StringComparison.Ordinal);
        var visualChanged = !string.Equals(
            left.VisualHash,
            right.VisualHash,
            StringComparison.Ordinal);

        if (!moved && !resized && !contentChanged && !visualChanged)
        {
            return null;
        }

        var descriptions = new List<string>();
        var label = GetElementLabel(left.Kind);
        if (contentChanged)
        {
            descriptions.Add(left.Kind == SlideElementKind.Image
                ? "image replaced"
                : $"{label.ToLowerInvariant()} content changed");
        }

        if (moved)
        {
            descriptions.Add($"{label.ToLowerInvariant()} moved");
        }

        if (resized)
        {
            descriptions.Add($"{label.ToLowerInvariant()} resized");
        }

        if (visualChanged && !contentChanged)
        {
            descriptions.Add(left.Kind == SlideElementKind.Image
                ? "image crop or formatting changed"
                : $"{label.ToLowerInvariant()} formatting changed");
        }

        var kind = contentChanged && left.Kind == SlideElementKind.Image
            ? SlideElementChangeKind.Replaced
            : resized
                ? SlideElementChangeKind.Resized
                : moved
                    ? SlideElementChangeKind.Moved
                    : SlideElementChangeKind.Modified;

        return new SlideElementChange(
            kind,
            left.Kind,
            string.Join(", ", descriptions),
            left.Bounds,
            right.Bounds);
    }

    private static double BoundsDistance(SlideBounds left, SlideBounds right)
    {
        var deltaX = (double)left.X - right.X;
        var deltaY = (double)left.Y - right.Y;
        var deltaWidth = (double)left.Width - right.Width;
        var deltaHeight = (double)left.Height - right.Height;
        return deltaX * deltaX + deltaY * deltaY +
               deltaWidth * deltaWidth + deltaHeight * deltaHeight;
    }

    private static string GetElementLabel(SlideElementKind kind) =>
        kind switch
        {
            SlideElementKind.Image => "Image",
            SlideElementKind.Text => "Text box",
            SlideElementKind.Shape => "Shape",
            SlideElementKind.Connector => "Connector",
            SlideElementKind.Group => "Shape group",
            SlideElementKind.Table => "Table",
            SlideElementKind.Chart => "Chart",
            _ => "Element"
        };

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
