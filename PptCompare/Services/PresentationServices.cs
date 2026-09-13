using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
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

public interface IConfigurablePresentationSourceService
{
    void UpdateLimits(PresentationReadLimits limits);
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

public sealed class OpenXmlPresentationSourceService : IPresentationSourceService, IConfigurablePresentationSourceService
{
    private static readonly XNamespace PresentationNamespace =
        "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace =
        "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private readonly Lock _limitsGate = new();
    private PresentationReadLimits _limits;

    public OpenXmlPresentationSourceService(PresentationReadLimits? limits = null)
    {
        _limits = limits ?? new PresentationReadLimits();
        ValidateLimits(_limits);
    }

    public void UpdateLimits(PresentationReadLimits limits)
    {
        ArgumentNullException.ThrowIfNull(limits);
        ValidateLimits(limits);
        lock (_limitsGate)
        {
            _limits = limits;
        }
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
            var limits = GetLimits();
            var path = ValidateFile(presentation.Location, limits);
            var openSettings = new OpenSettings
            {
                AutoSave = false,
                MaxCharactersInPart = limits.MaxCharactersPerPart
            };

            using var document = PresentationDocument.Open(path, false, openSettings);
            ValidateDocumentType(document.DocumentType, Path.GetExtension(path));
            var presentationPart = document.PresentationPart
                ?? throw new PresentationLoadException(
                    "The selected file does not contain a readable PowerPoint presentation.");
            var presentationRoot = presentationPart.Presentation
                ?? throw new PresentationLoadException(
                    "The selected file does not contain a readable presentation root.");
            var slideIdList = presentationRoot.SlideIdList
                ?? throw new PresentationLoadException("The selected presentation does not contain any slides.");
            var slideIds = slideIdList.Elements<P.SlideId>().Take(limits.MaxSlides + 1).ToList();

            if (slideIds.Count > limits.MaxSlides)
            {
                throw new PresentationLoadException(
                    $"The presentation contains more than the supported limit of {limits.MaxSlides:N0} slides.");
            }

            var slides = new List<PresentationSlide>();
            long totalCharacters = 0;
            var relatedPartContext = new RelatedPartReadContext(limits);
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
                    .Take(limits.MaxParagraphsPerSlide + 1)
                    .Count();
                if (paragraphCount > limits.MaxParagraphsPerSlide)
                {
                    throw new PresentationLoadException(
                        $"Slide {slides.Count + 1} contains too many text paragraphs to process safely.");
                }

                var drawableElementCount = slideXml
                    .Descendants(PresentationNamespace + "spTree")
                    .FirstOrDefault()?
                    .Elements()
                    .Where(IsDrawableElement)
                    .Take(limits.MaxElementsPerSlide + 1)
                    .Count() ?? 0;
                if (drawableElementCount > limits.MaxElementsPerSlide)
                {
                    throw new PresentationLoadException(
                        $"Slide {slides.Count + 1} contains too many elements to process safely.");
                }

                var tableCellCount = slideXml
                    .Descendants(DrawingNamespace + "tc")
                    .Take(limits.MaxTableCellsPerSlide + 1)
                    .Count();
                if (tableCellCount > limits.MaxTableCellsPerSlide)
                {
                    throw new PresentationLoadException(
                        $"Slide {slides.Count + 1} contains too many table cells to process safely.");
                }

                var (paragraphs, titleContent, textContent) =
                    ExtractTextInReadingOrder(slidePart, slideXml);
                var slideCharacters = 0;
                foreach (var text in paragraphs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    slideCharacters += text.Length;
                    totalCharacters += text.Length;
                    if (slideCharacters > limits.MaxCharactersPerSlide ||
                        totalCharacters > limits.MaxTotalCharacters)
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
                    limits,
                    relatedPartContext,
                    cancellationToken);
                slides.Add(new PresentationSlide(
                    slides.Count + 1,
                    paragraphs[0],
                    paragraphs,
                    slideWidth,
                    slideHeight,
                    elements)
                {
                    TitleContent = titleContent,
                    TextContent = textContent
                });
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

    private PresentationReadLimits GetLimits()
    {
        lock (_limitsGate)
        {
            return _limits;
        }
    }

    private static string ValidateFile(string location, PresentationReadLimits limits)
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

        if (file.Length > limits.MaxFileBytes)
        {
            throw new PresentationLoadException(
                $"The selected presentation is larger than the supported limit of {limits.MaxFileBytes / 1024 / 1024:N0} MB.");
        }

        return fullPath;
    }

    private static void ValidateDocumentType(
        PresentationDocumentType documentType,
        string extension)
    {
        var expectedType = extension.Equals(".pptm", StringComparison.OrdinalIgnoreCase)
            ? PresentationDocumentType.MacroEnabledPresentation
            : PresentationDocumentType.Presentation;
        if (documentType != expectedType)
        {
            throw new PresentationLoadException(
                "The selected file's PowerPoint document type does not match its file extension.");
        }
    }

    private static void ValidateLimits(PresentationReadLimits limits)
    {
        if (limits.MaxFileBytes <= 0 ||
            limits.MaxCharactersPerPart <= 0 ||
            limits.MaxSlides <= 0 ||
            limits.MaxElementsPerSlide <= 0 ||
            limits.MaxTableCellsPerSlide <= 0 ||
            limits.MaxParagraphsPerSlide <= 0 ||
            limits.MaxCharactersPerSlide <= 0 ||
            limits.MaxTotalCharacters <= 0 ||
            limits.MaxRelatedPartBytes <= 0 ||
            limits.MaxRelationshipReferences <= 0 ||
            limits.MaxUniqueRelatedParts <= 0 ||
            limits.MaxTotalRelatedPartBytes <= 0 ||
            limits.MaxPreviewImageBytes <= 0 ||
            limits.MaxTotalPreviewImageBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(limits), "All presentation read limits must be positive.");
        }

        if (limits.MaxUniqueRelatedParts > limits.MaxRelationshipReferences ||
            limits.MaxTotalRelatedPartBytes < limits.MaxRelatedPartBytes ||
            limits.MaxTotalPreviewImageBytes < limits.MaxPreviewImageBytes)
        {
            throw new ArgumentOutOfRangeException(
                nameof(limits),
                "Presentation-wide limits must be consistent with their per-item limits.");
        }
    }

    private static ExtractedSlideText ExtractTextInReadingOrder(
        SlidePart slidePart,
        XElement slideXml)
    {
        var themeColors = ReadThemeColors(slidePart);
        var shapeTree = slideXml.Descendants(PresentationNamespace + "spTree").FirstOrDefault();
        if (shapeTree is null)
        {
            return CreateFallbackText(slideXml, themeColors);
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
                ExtractTextBlocks(element, themeColors)))
            .Where(candidate => candidate.Paragraphs.Count > 0)
            .ToList();
        if (candidates.Count == 0)
        {
            return CreateFallbackText(slideXml, themeColors);
        }

        // PowerPoint stores shapes in drawing/z-order, which is not necessarily reading order.
        // The first entry is metadata for the comparison heading. Keep every actual paragraph,
        // including that title text at its visible position, in the spatially ordered body.
        var titleCandidate = candidates.FirstOrDefault(candidate => candidate.IsTitle) ?? candidates[0];
        var title = titleCandidate.Paragraphs[0];
        var blocks = candidates
            .OrderBy(candidate => candidate.HasResolvedBounds ? 0 : 1)
            .ThenBy(candidate => candidate.Bounds.Y)
            .ThenBy(candidate => candidate.Bounds.X)
            .ThenBy(candidate => candidate.DocumentOrder)
            .SelectMany(candidate => candidate.Blocks)
            .ToList();
        var paragraphs = new List<string> { title.Text };
        paragraphs.AddRange(EnumerateParagraphs(blocks).Select(paragraph => paragraph.Text));
        return new ExtractedSlideText(paragraphs, title, blocks);
    }

    private static ExtractedSlideText CreateFallbackText(
        XElement element,
        IReadOnlyDictionary<string, string> themeColors)
    {
        var richParagraphs = element
            .Descendants(DrawingNamespace + "p")
            .Select(paragraph => ExtractTextParagraph(paragraph, themeColors))
            .Where(paragraph => paragraph.Text.Length > 0)
            .ToList();
        var title = richParagraphs.FirstOrDefault();
        List<SlideTextBlock> blocks =
            [.. richParagraphs.Select(paragraph => new SlideTextParagraphBlock(paragraph))];
        var plainParagraphs = richParagraphs.Select(paragraph => paragraph.Text).ToList();
        return new ExtractedSlideText(plainParagraphs, title, blocks);
    }

    private static List<SlideTextBlock> ExtractTextBlocks(
        XElement element,
        IReadOnlyDictionary<string, string> themeColors)
    {
        var table = element.Descendants(DrawingNamespace + "tbl").FirstOrDefault();
        if (table is not null)
        {
            return [ExtractTextTable(table, themeColors)];
        }

        return
        [
            .. element
                .Descendants(DrawingNamespace + "p")
                .Select(paragraph => ExtractTextParagraph(paragraph, themeColors))
                .Where(paragraph => paragraph.Text.Length > 0)
                .Select(paragraph => new SlideTextParagraphBlock(paragraph))
        ];
    }

    private static SlideTextTableBlock ExtractTextTable(
        XElement table,
        IReadOnlyDictionary<string, string> themeColors)
    {
        List<SlideTextTableRow> rows =
        [
            .. table
                .Elements(DrawingNamespace + "tr")
                .Select(row => new SlideTextTableRow(
                [
                    .. row
                        .Elements(DrawingNamespace + "tc")
                        .Select(cell => new SlideTextTableCell(
                        [
                            .. cell
                                .Descendants(DrawingNamespace + "p")
                                .Select(paragraph => ExtractTextParagraph(paragraph, themeColors))
                                .Where(paragraph => paragraph.Text.Length > 0)
                        ],
                        ReadSpan(cell, "gridSpan"),
                        ReadSpan(cell, "rowSpan")))
                ]))
        ];
        var gridColumns = table
            .Element(DrawingNamespace + "tblGrid")?
            .Elements(DrawingNamespace + "gridCol")
            .Count() ?? 0;
        var populatedColumns = rows.Count == 0
            ? 0
            : rows.Max(row => row.Cells.Sum(cell => Math.Max(1, cell.ColumnSpan)));
        return new SlideTextTableBlock(rows, Math.Max(1, Math.Max(gridColumns, populatedColumns)));
    }

    private static int ReadSpan(XElement cell, string attributeName)
    {
        var rawValue = cell.Attribute(attributeName)?.Value ??
                       cell.Element(DrawingNamespace + "tcPr")?.Attribute(attributeName)?.Value;
        return int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var span)
            ? Math.Clamp(span, 1, 100)
            : 1;
    }

    private static IEnumerable<SlideTextParagraph> EnumerateParagraphs(
        IEnumerable<SlideTextBlock> blocks)
    {
        foreach (var block in blocks)
        {
            if (block is SlideTextParagraphBlock paragraphBlock)
            {
                if (paragraphBlock.Paragraph.Text.Length > 0)
                {
                    yield return paragraphBlock.Paragraph;
                }

                continue;
            }

            if (block is not SlideTextTableBlock table)
            {
                continue;
            }

            foreach (var paragraph in table.Rows
                         .SelectMany(row => row.Cells)
                         .SelectMany(cell => cell.Paragraphs)
                         .Where(paragraph => paragraph.Text.Length > 0))
            {
                yield return paragraph;
            }
        }
    }

    private static SlideTextParagraph ExtractTextParagraph(
        XElement paragraph,
        IReadOnlyDictionary<string, string> themeColors)
    {
        var paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
        var level = int.TryParse(
            paragraphProperties?.Attribute("lvl")?.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out var parsedLevel)
            ? Math.Clamp(parsedLevel + 1, 1, 9)
            : 1;
        var listDefault = paragraph
            .Ancestors()
            .FirstOrDefault(value => value.Name.LocalName == "txBody")?
            .Element(DrawingNamespace + "lstStyle")?
            .Element(DrawingNamespace + $"lvl{level}pPr")?
            .Element(DrawingNamespace + "defRPr");
        var paragraphDefault = paragraphProperties?.Element(DrawingNamespace + "defRPr");
        var runs = new List<SlideTextRun>();

        foreach (var child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                runs.Add(new SlideTextRun(
                    Environment.NewLine,
                    ReadTextStyle(
                        themeColors,
                        listDefault,
                        paragraphDefault,
                        child.Element(DrawingNamespace + "rPr"))));
                continue;
            }

            if (child.Name != DrawingNamespace + "r" && child.Name != DrawingNamespace + "fld")
            {
                continue;
            }

            var text = child.Element(DrawingNamespace + "t")?.Value;
            if (!string.IsNullOrEmpty(text))
            {
                runs.Add(new SlideTextRun(
                    text,
                    ReadTextStyle(
                        themeColors,
                        listDefault,
                        paragraphDefault,
                        child.Element(DrawingNamespace + "rPr"))));
            }
        }

        if (runs.Count == 0)
        {
            var text = string.Concat(
                paragraph.Descendants(DrawingNamespace + "t").Select(value => value.Value));
            if (text.Length > 0)
            {
                runs.Add(new SlideTextRun(
                    text,
                    ReadTextStyle(themeColors, listDefault, paragraphDefault)));
            }
        }

        TrimRuns(runs);
        return new SlideTextParagraph(runs);
    }

    private static void TrimRuns(List<SlideTextRun> runs)
    {
        while (runs.Count > 0)
        {
            var trimmed = runs[0].Text.TrimStart();
            if (trimmed.Length > 0)
            {
                runs[0] = runs[0] with { Text = trimmed };
                break;
            }

            runs.RemoveAt(0);
        }

        while (runs.Count > 0)
        {
            var last = runs.Count - 1;
            var trimmed = runs[last].Text.TrimEnd();
            if (trimmed.Length > 0)
            {
                runs[last] = runs[last] with { Text = trimmed };
                break;
            }

            runs.RemoveAt(last);
        }
    }

    private static SlideTextStyle ReadTextStyle(
        IReadOnlyDictionary<string, string> themeColors,
        params XElement?[] layers)
    {
        bool? bold = null;
        bool? italic = null;
        bool? underline = null;
        bool? strikeThrough = null;
        var baseline = SlideTextBaseline.Normal;
        string? color = null;
        string? fontFamily = null;

        foreach (var layer in layers.Where(layer => layer is not null))
        {
            bold = ReadOptionalBoolean(layer!, "b") ?? bold;
            italic = ReadOptionalBoolean(layer!, "i") ?? italic;
            var underlineValue = layer!.Attribute("u")?.Value;
            if (underlineValue is not null)
            {
                underline = !string.Equals(underlineValue, "none", StringComparison.OrdinalIgnoreCase);
            }

            var strikeValue = layer.Attribute("strike")?.Value;
            if (strikeValue is not null)
            {
                strikeThrough = !string.Equals(strikeValue, "noStrike", StringComparison.OrdinalIgnoreCase);
            }

            if (int.TryParse(
                    layer.Attribute("baseline")?.Value,
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out var baselineValue))
            {
                baseline = baselineValue switch
                {
                    > 0 => SlideTextBaseline.Superscript,
                    < 0 => SlideTextBaseline.Subscript,
                    _ => SlideTextBaseline.Normal
                };
            }

            var typeface = layer.Element(DrawingNamespace + "latin")?.Attribute("typeface")?.Value;
            if (!string.IsNullOrWhiteSpace(typeface) && !typeface.StartsWith('+'))
            {
                fontFamily = typeface;
            }

            color = ReadTextColor(layer, themeColors) ?? color;
        }

        return new SlideTextStyle(
            bold,
            italic,
            underline,
            strikeThrough,
            baseline,
            color,
            fontFamily);
    }

    private static bool? ReadOptionalBoolean(XElement element, string attributeName)
    {
        var value = element.Attribute(attributeName)?.Value;
        if (value is null)
        {
            return null;
        }

        return value is "1" or "true" or "on";
    }

    private static string? ReadTextColor(
        XElement properties,
        IReadOnlyDictionary<string, string> themeColors)
    {
        var color = properties.Element(DrawingNamespace + "solidFill")?.Elements().FirstOrDefault();
        if (color is null)
        {
            return null;
        }

        var rawValue = color.Name.LocalName switch
        {
            "srgbClr" => color.Attribute("val")?.Value,
            "sysClr" => color.Attribute("lastClr")?.Value,
            "schemeClr" when color.Attribute("val")?.Value is { } scheme &&
                             themeColors.TryGetValue(scheme, out var themeColor) => themeColor,
            _ => null
        };
        return IsRgbHex(rawValue) ? $"#{rawValue}" : null;
    }

    private static Dictionary<string, string> ReadThemeColors(SlidePart slidePart)
    {
        var theme = slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme;
        if (theme is null)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var themeXml = XElement.Parse(theme.OuterXml, LoadOptions.None);
        var colorScheme = themeXml.Descendants(DrawingNamespace + "clrScheme").FirstOrDefault();
        if (colorScheme is null)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var colors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in colorScheme.Elements())
        {
            var color = entry.Elements().FirstOrDefault();
            var value = color?.Name.LocalName switch
            {
                "srgbClr" => color.Attribute("val")?.Value,
                "sysClr" => color.Attribute("lastClr")?.Value,
                _ => null
            };
            if (IsRgbHex(value))
            {
                colors[entry.Name.LocalName] = value!;
            }
        }

        return colors;
    }

    private static bool IsRgbHex(string? value) =>
        value is { Length: 6 } && value.All(Uri.IsHexDigit);

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
            var inheritedElement = inheritedRoot?
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

    private static bool IsTitlePlaceholder(XElement element)
    {
        var placeholderType = element
            .Descendants(PresentationNamespace + "ph")
            .FirstOrDefault()?
            .Attribute("type")?
            .Value;
        return placeholderType is "title" or "ctrTitle";
    }

    private static List<SlideElement> ExtractElements(
        SlidePart slidePart,
        XElement slideXml,
        PresentationReadLimits limits,
        RelatedPartReadContext relatedPartContext,
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
                limits,
                relatedPartContext,
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

    private static RelatedContent ReadRelatedContent(
        SlidePart slidePart,
        XElement element,
        bool capturePreview,
        PresentationReadLimits limits,
        RelatedPartReadContext context,
        CancellationToken cancellationToken)
    {
        var relationshipIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var attribute in element.DescendantsAndSelf().Attributes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (attribute.Name.Namespace != RelationshipNamespace ||
                string.IsNullOrWhiteSpace(attribute.Value) ||
                !relationshipIds.Add(attribute.Value))
            {
                continue;
            }

            context.RecordRelationshipReference();
        }

        var hashes = new List<string>();
        byte[]? previewBytes = null;

        foreach (var relationshipId in relationshipIds.Order(StringComparer.Ordinal))
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

            var partUri = part.Uri;
            if (!context.TryGet(partUri, out var partContent))
            {
                context.RecordUniquePart();
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                partContent = ReadAndHashPart(
                    stream,
                    part is ImagePart,
                    limits,
                    context,
                    cancellationToken);
                context.Add(partUri, partContent);
            }

            hashes.Add(partContent.Hash);
            if (capturePreview && previewBytes is null && part is ImagePart)
            {
                previewBytes = partContent.PreviewBytes;
            }
        }

        return new RelatedContent(
            hashes.Count == 0 ? string.Empty : HashText(string.Join("|", hashes)),
            previewBytes);
    }

    private static RelatedContent ReadAndHashPart(
        Stream stream,
        bool capturePreview,
        PresentationReadLimits limits,
        RelatedPartReadContext context,
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

            try
            {
                totalBytes = checked(totalBytes + read);
            }
            catch (OverflowException exception)
            {
                throw new PresentationLoadException(
                    "The presentation contains an embedded item that is too large to process safely.",
                    exception);
            }

            if (totalBytes > limits.MaxRelatedPartBytes)
            {
                throw new PresentationLoadException(
                    "The presentation contains an embedded item that is too large to process safely.");
            }

            context.RecordProcessedBytes(read);

            hash.AppendData(buffer, 0, read);
            if (!keepPreview)
            {
                continue;
            }

            if (totalBytes > limits.MaxPreviewImageBytes ||
                !context.CanRetainPreview(totalBytes))
            {
                keepPreview = false;
                continue;
            }

            preview!.Write(buffer, 0, read);
        }

        var previewBytes = keepPreview ? preview?.ToArray() : null;
        if (previewBytes is not null)
        {
            context.RecordPreviewBytes(previewBytes.Length);
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
        IReadOnlyList<SlideTextBlock> Blocks)
    {
        public bool HasResolvedBounds => Bounds.Width > 0 || Bounds.Height > 0;
        public List<SlideTextParagraph> Paragraphs =>
            [.. EnumerateParagraphs(Blocks)];
    }

    private sealed record ExtractedSlideText(
        List<string> PlainParagraphs,
        SlideTextParagraph? Title,
        IReadOnlyList<SlideTextBlock> Blocks);

    private sealed record RelatedContent(string Hash, byte[]? PreviewBytes);

    private sealed class RelatedPartReadContext(PresentationReadLimits limits)
    {
        private readonly Dictionary<Uri, RelatedContent> _parts = [];
        private int _relationshipReferences;
        private int _uniqueParts;
        private long _processedBytes;
        private long _previewBytes;

        public bool TryGet(Uri partUri, out RelatedContent content) =>
            _parts.TryGetValue(partUri, out content!);

        public void Add(Uri partUri, RelatedContent content) =>
            _parts.Add(partUri, content);

        public void RecordRelationshipReference()
        {
            try
            {
                _relationshipReferences = checked(_relationshipReferences + 1);
            }
            catch (OverflowException exception)
            {
                throw CreateRelationshipLimitException(exception);
            }

            if (_relationshipReferences > limits.MaxRelationshipReferences)
            {
                throw CreateRelationshipLimitException();
            }
        }

        public void RecordUniquePart()
        {
            try
            {
                _uniqueParts = checked(_uniqueParts + 1);
            }
            catch (OverflowException exception)
            {
                throw CreateRelationshipLimitException(exception);
            }

            if (_uniqueParts > limits.MaxUniqueRelatedParts)
            {
                throw CreateRelationshipLimitException();
            }
        }

        public void RecordProcessedBytes(int bytesRead)
        {
            try
            {
                _processedBytes = checked(_processedBytes + bytesRead);
            }
            catch (OverflowException exception)
            {
                throw CreateEmbeddedContentLimitException(exception);
            }

            if (_processedBytes > limits.MaxTotalRelatedPartBytes)
            {
                throw CreateEmbeddedContentLimitException();
            }
        }

        public bool CanRetainPreview(long currentPartBytes) =>
            currentPartBytes <= limits.MaxTotalPreviewImageBytes - _previewBytes;

        public void RecordPreviewBytes(int previewBytes)
        {
            try
            {
                _previewBytes = checked(_previewBytes + previewBytes);
            }
            catch (OverflowException exception)
            {
                throw CreateEmbeddedContentLimitException(exception);
            }

            if (_previewBytes > limits.MaxTotalPreviewImageBytes)
            {
                throw CreateEmbeddedContentLimitException();
            }
        }

        private static PresentationLoadException CreateRelationshipLimitException(
            Exception? innerException = null) =>
            innerException is null
                ? new PresentationLoadException(
                    "The presentation contains too many embedded-content references to process safely.")
                : new PresentationLoadException(
                    "The presentation contains too many embedded-content references to process safely.",
                    innerException);

        private static PresentationLoadException CreateEmbeddedContentLimitException(
            Exception? innerException = null) =>
            innerException is null
                ? new PresentationLoadException(
                    "The presentation contains too much embedded content to process safely.")
                : new PresentationLoadException(
                    "The presentation contains too much embedded content to process safely.",
                    innerException);
    }
}

public sealed partial class TextPresentationComparisonService : IPresentationComparisonService
{
    private const int MaxDiffTokensPerSide = 4_000;
    private const long MaxElementCandidateChecksPerComparison = 2_000_000;

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
            var remainingElementCandidateChecks = MaxElementCandidateChecksPerComparison;

            for (var rightIndex = 0; rightIndex < right.Slides.Count; rightIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!matches.TryGetValue(rightIndex, out var leftIndex))
                {
                    results.Add(CreateComparison(
                        rightIndex + 1,
                        null,
                        right.Slides[rightIndex],
                        ref remainingElementCandidateChecks,
                        cancellationToken));
                    continue;
                }

                matchedLeft.Add(leftIndex);
                var comparison = CreateComparison(
                    rightIndex + 1,
                    left.Slides[leftIndex],
                    right.Slides[rightIndex],
                    ref remainingElementCandidateChecks,
                    cancellationToken);
                results.Add(leftIndex == rightIndex
                    ? comparison
                    : MarkAsMoved(comparison, leftIndex + 1, rightIndex + 1));
            }

            for (var leftIndex = 0; leftIndex < left.Slides.Count; leftIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!matchedLeft.Contains(leftIndex))
                {
                    results.Add(CreateComparison(
                        results.Count + 1,
                        left.Slides[leftIndex],
                        null,
                        ref remainingElementCandidateChecks,
                        cancellationToken));
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
        var unmatchedLeft = new SortedSet<int>(Enumerable.Range(0, leftSlides.Count));
        var leftFingerprints = new List<string>(leftSlides.Count);
        foreach (var slide in leftSlides)
        {
            cancellationToken.ThrowIfCancellationRequested();
            leftFingerprints.Add(CreateSlideFingerprint(slide, cancellationToken));
        }

        var rightFingerprints = new List<string>(rightSlides.Count);
        foreach (var slide in rightSlides)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rightFingerprints.Add(CreateSlideFingerprint(slide, cancellationToken));
        }

        var leftByFingerprint = new Dictionary<string, SortedSet<int>>(StringComparer.Ordinal);
        for (var leftIndex = 0; leftIndex < leftFingerprints.Count; leftIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            AddCandidate(leftByFingerprint, leftFingerprints[leftIndex], leftIndex);
        }

        // Match identical slides first. This identifies pure reordering without
        // confusing it with edits to slides that happen to occupy the same position.
        for (var rightIndex = 0; rightIndex < rightSlides.Count; rightIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!leftByFingerprint.TryGetValue(
                    rightFingerprints[rightIndex],
                    out var exactMatches) ||
                exactMatches.Count == 0)
            {
                continue;
            }

            var leftMatch = TakeNearest(exactMatches, rightIndex);
            matches[rightIndex] = leftMatch;
            unmatchedLeft.Remove(leftMatch);
        }

        var leftByTitle = new Dictionary<string, SortedSet<int>>(StringComparer.Ordinal);
        foreach (var leftIndex in unmatchedLeft)
        {
            cancellationToken.ThrowIfCancellationRequested();
            AddCandidate(leftByTitle, Normalize(leftSlides[leftIndex].Title), leftIndex);
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
            var leftMatch = leftByTitle.TryGetValue(normalizedTitle, out var sameTitle) &&
                            sameTitle.Count > 0
                ? TakeNearest(sameTitle, rightIndex)
                : unmatchedLeft.Contains(rightIndex)
                    ? rightIndex
                    : FindNearest(unmatchedLeft, rightIndex);
            matches[rightIndex] = leftMatch;
            unmatchedLeft.Remove(leftMatch);
            var matchedTitle = Normalize(leftSlides[leftMatch].Title);
            if (leftByTitle.TryGetValue(matchedTitle, out var titleMatches))
            {
                titleMatches.Remove(leftMatch);
            }
        }

        return matches;
    }

    private static void AddCandidate(
        Dictionary<string, SortedSet<int>> index,
        string key,
        int candidate)
    {
        if (!index.TryGetValue(key, out var candidates))
        {
            candidates = [];
            index.Add(key, candidates);
        }

        candidates.Add(candidate);
    }

    private static int TakeNearest(SortedSet<int> candidates, int expectedIndex)
    {
        var nearest = FindNearest(candidates, expectedIndex);
        candidates.Remove(nearest);
        return nearest;
    }

    private static int FindNearest(SortedSet<int> candidates, int expectedIndex)
    {
        if (candidates.Contains(expectedIndex))
        {
            return expectedIndex;
        }

        var lowerView = candidates.GetViewBetween(int.MinValue, expectedIndex);
        var upperView = candidates.GetViewBetween(expectedIndex, int.MaxValue);
        var hasLower = TryGetLast(lowerView, out var lower);
        var hasUpper = TryGetFirst(upperView, out var upper);
        if (!hasLower)
        {
            return upper;
        }

        if (!hasUpper)
        {
            return lower;
        }

        return expectedIndex - lower <= upper - expectedIndex ? lower : upper;
    }

    private static bool TryGetFirst(SortedSet<int> candidates, out int value)
    {
        foreach (var candidate in candidates)
        {
            value = candidate;
            return true;
        }

        value = 0;
        return false;
    }

    private static bool TryGetLast(SortedSet<int> candidates, out int value)
    {
        foreach (var candidate in candidates.Reverse())
        {
            value = candidate;
            return true;
        }

        value = 0;
        return false;
    }

    private static string CreateSlideFingerprint(
        PresentationSlide slide,
        CancellationToken cancellationToken)
    {
        var value = new StringBuilder();
        value.Append(Normalize(slide.Title));
        foreach (var paragraph in slide.Paragraphs.Skip(1))
        {
            cancellationToken.ThrowIfCancellationRequested();
            value.Append('|').Append(Normalize(paragraph));
        }

        foreach (var element in slide.Elements
                     .OrderBy(element => element.Kind)
                     .ThenBy(element => element.Bounds.Y)
                     .ThenBy(element => element.Bounds.X)
                     .ThenBy(element => element.ContentHash, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
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
            LeftTitleContent = item.LeftTitleContent,
            LeftBodyContent = item.LeftBodyContent,
            RightTitleContent = item.RightTitleContent,
            RightBodyContent = item.RightBodyContent,
            LeftSlide = item.LeftSlide,
            RightSlide = item.RightSlide,
            ElementChanges = item.ElementChanges
        };
    }

    private static SlideComparisonItem CreateComparison(
        int number,
        PresentationSlide? left,
        PresentationSlide? right,
        ref long remainingElementCandidateChecks,
        CancellationToken cancellationToken)
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
                RightTitleContent = CreateTitleContent(right),
                RightBodyContent = right.TextContent,
                RightSlide = right,
                ElementChanges = CreateUnmatchedElementChanges(
                    right.Elements,
                    SlideElementChangeKind.Added,
                    cancellationToken)
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
                LeftTitleContent = CreateTitleContent(left),
                LeftBodyContent = left.TextContent,
                LeftSlide = left,
                ElementChanges = CreateUnmatchedElementChanges(
                    left.Elements,
                    SlideElementChangeKind.Removed,
                    cancellationToken)
            };
        }

        var leftBody = FormatBody(left);
        var rightBody = FormatBody(right);
        var titleDiff = BuildWordDiff(left.Title, right.Title, cancellationToken);
        var bodyDiff = BuildParagraphDiff(
            [.. left.Paragraphs.Skip(1)],
            [.. right.Paragraphs.Skip(1)],
            cancellationToken);
        var titleChanged = !string.Equals(
            Normalize(left.Title),
            Normalize(right.Title),
            StringComparison.Ordinal);
        var changedBlockCount = bodyDiff.ChangedBlocks + (titleChanged ? 1 : 0);
        var elementChanges = CompareElements(
            left.Elements,
            right.Elements,
            ref remainingElementCandidateChecks,
            cancellationToken);
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
            LeftTitleContent = CreateTitleContent(left),
            LeftBodyContent = left.TextContent,
            RightTitleContent = CreateTitleContent(right),
            RightBodyContent = right.TextContent,
            LeftSlide = left,
            RightSlide = right,
            ElementChanges = elementChanges
        };
    }

    private static string FormatBody(PresentationSlide slide) =>
        string.Join(Environment.NewLine + Environment.NewLine, slide.Paragraphs.Skip(1));

    private static IReadOnlyList<SlideTextBlock> CreateTitleContent(PresentationSlide slide) =>
        slide.TitleContent is null
            ? []
            : [new SlideTextParagraphBlock(slide.TitleContent)];

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

    private static List<SlideElementChange> CreateUnmatchedElementChanges(
        IReadOnlyList<SlideElement> elements,
        SlideElementChangeKind changeKind,
        CancellationToken cancellationToken)
    {
        var changes = new List<SlideElementChange>(elements.Count);
        foreach (var element in elements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var added = changeKind == SlideElementChangeKind.Added;
            changes.Add(new SlideElementChange(
                changeKind,
                element.Kind,
                $"{GetElementLabel(element.Kind)} {(added ? "added" : "removed")}",
                added ? null : element.Bounds,
                added ? element.Bounds : null));
        }

        return changes;
    }

    private static List<SlideElementChange> CompareElements(
        IReadOnlyList<SlideElement> leftElements,
        IReadOnlyList<SlideElement> rightElements,
        ref long remainingCandidateChecks,
        CancellationToken cancellationToken)
    {
        var changes = new List<SlideElementChange>();
        var candidates = new ElementCandidateIndex(rightElements, cancellationToken);

        foreach (var leftElement in leftElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rightIndex = candidates.TakeBestMatch(
                leftElement,
                ref remainingCandidateChecks,
                cancellationToken);
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

            var rightElement = rightElements[rightIndex];
            var change = CompareMatchedElement(leftElement, rightElement);
            if (change is not null)
            {
                changes.Add(change);
            }
        }

        foreach (var rightIndex in candidates.EnumerateUnmatched())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var element = rightElements[rightIndex];
            changes.Add(new SlideElementChange(
                SlideElementChangeKind.Added,
                element.Kind,
                $"{GetElementLabel(element.Kind)} added",
                null,
                element.Bounds));
        }

        return changes;
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

    private readonly record struct ElementLookupKey(
        SlideElementKind Kind,
        string Value);

    private sealed class ElementLookupKeyComparer(StringComparer valueComparer)
        : IEqualityComparer<ElementLookupKey>
    {
        public bool Equals(ElementLookupKey left, ElementLookupKey right) =>
            left.Kind == right.Kind && valueComparer.Equals(left.Value, right.Value);

        public int GetHashCode(ElementLookupKey key) =>
            HashCode.Combine(key.Kind, valueComparer.GetHashCode(key.Value));
    }

    private sealed class ElementCandidateIndex
    {
        private readonly IReadOnlyList<SlideElement> _elements;
        private readonly bool[] _unmatched;
        private readonly Dictionary<ElementLookupKey, SortedSet<int>> _byId =
            new(new ElementLookupKeyComparer(StringComparer.Ordinal));
        private readonly Dictionary<ElementLookupKey, SortedSet<int>> _byContent =
            new(new ElementLookupKeyComparer(StringComparer.Ordinal));
        private readonly Dictionary<ElementLookupKey, SortedSet<int>> _byName =
            new(new ElementLookupKeyComparer(StringComparer.OrdinalIgnoreCase));
        private readonly Dictionary<SlideElementKind, SortedSet<int>> _byKind = [];

        public ElementCandidateIndex(
            IReadOnlyList<SlideElement> elements,
            CancellationToken cancellationToken)
        {
            _elements = elements;
            _unmatched = new bool[elements.Count];
            for (var index = 0; index < elements.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var element = elements[index];
                _unmatched[index] = true;
                Add(_byId, new ElementLookupKey(element.Kind, element.Id), index);
                Add(_byContent, new ElementLookupKey(element.Kind, element.ContentHash), index);
                Add(
                    _byName,
                    new ElementLookupKey(element.Kind, element.Name),
                    index);
                Add(_byKind, element.Kind, index);
            }
        }

        public int TakeBestMatch(
            SlideElement left,
            ref long remainingCandidateChecks,
            CancellationToken cancellationToken)
        {
            if (TryTakeAny(
                    _byId,
                    new ElementLookupKey(left.Kind, left.Id),
                    out var match))
            {
                return match;
            }

            if (TryTakeClosest(
                    _byContent,
                    new ElementLookupKey(left.Kind, left.ContentHash),
                    left.Bounds,
                    ref remainingCandidateChecks,
                    cancellationToken,
                    out match) ||
                TryTakeClosest(
                    _byName,
                    new ElementLookupKey(left.Kind, left.Name),
                    left.Bounds,
                    ref remainingCandidateChecks,
                    cancellationToken,
                    out match) ||
                TryTakeClosest(
                    _byKind,
                    left.Kind,
                    left.Bounds,
                    ref remainingCandidateChecks,
                    cancellationToken,
                    out match))
            {
                return match;
            }

            return -1;
        }

        public IEnumerable<int> EnumerateUnmatched()
        {
            for (var index = 0; index < _unmatched.Length; index++)
            {
                if (_unmatched[index])
                {
                    yield return index;
                }
            }
        }

        private static void Add<TKey>(
            Dictionary<TKey, SortedSet<int>> index,
            TKey key,
            int elementIndex)
            where TKey : notnull
        {
            if (!index.TryGetValue(key, out var matches))
            {
                matches = [];
                index.Add(key, matches);
            }

            matches.Add(elementIndex);
        }

        private bool TryTakeAny<TKey>(
            IReadOnlyDictionary<TKey, SortedSet<int>> index,
            TKey key,
            out int match)
            where TKey : notnull
        {
            if (!index.TryGetValue(key, out var candidates) || candidates.Count == 0)
            {
                match = -1;
                return false;
            }

            match = candidates.Min;
            Remove(match);
            return true;
        }

        private bool TryTakeClosest<TKey>(
            IReadOnlyDictionary<TKey, SortedSet<int>> index,
            TKey key,
            SlideBounds expectedBounds,
            ref long remainingCandidateChecks,
            CancellationToken cancellationToken,
            out int match)
            where TKey : notnull
        {
            if (!index.TryGetValue(key, out var candidates) || candidates.Count == 0)
            {
                match = -1;
                return false;
            }

            if (remainingCandidateChecks <= 0 ||
                candidates.Count > remainingCandidateChecks)
            {
                match = -1;
                return false;
            }

            match = -1;
            var bestDistance = double.MaxValue;
            foreach (var candidate in candidates)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var distance = BoundsDistance(expectedBounds, _elements[candidate].Bounds);
                if (distance < bestDistance)
                {
                    match = candidate;
                    bestDistance = distance;
                }

                remainingCandidateChecks--;
            }

            Remove(match);
            return true;
        }

        private void Remove(int index)
        {
            var element = _elements[index];
            _unmatched[index] = false;
            Remove(_byId, new ElementLookupKey(element.Kind, element.Id), index);
            Remove(_byContent, new ElementLookupKey(element.Kind, element.ContentHash), index);
            Remove(
                _byName,
                new ElementLookupKey(element.Kind, element.Name),
                index);
            Remove(_byKind, element.Kind, index);
        }

        private static void Remove<TKey>(
            IReadOnlyDictionary<TKey, SortedSet<int>> index,
            TKey key,
            int elementIndex)
            where TKey : notnull
        {
            if (index.TryGetValue(key, out var matches))
            {
                matches.Remove(elementIndex);
            }
        }
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
            List<string> rightParagraphs,
            CancellationToken cancellationToken)
    {
        var matches = FindMatchingParagraphs(
            leftParagraphs,
            rightParagraphs,
            cancellationToken);
        var leftSegments = new List<DiffSegment>();
        var rightSegments = new List<DiffSegment>();
        var changedBlocks = 0;
        var leftStart = 0;
        var rightStart = 0;

        foreach (var match in matches.Append((
                     Left: leftParagraphs.Count,
                     Right: rightParagraphs.Count)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var leftCount = match.Left - leftStart;
            var rightCount = match.Right - rightStart;
            var pairedCount = Math.Min(leftCount, rightCount);

            for (var offset = 0; offset < pairedCount; offset++)
            {
                var wordDiff = BuildWordDiff(
                    leftParagraphs[leftStart + offset],
                    rightParagraphs[rightStart + offset],
                    cancellationToken);
                AppendParagraph(leftSegments, wordDiff.Left);
                AppendParagraph(rightSegments, wordDiff.Right);
            }

            for (var offset = pairedCount; offset < leftCount; offset++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                AppendParagraph(
                    leftSegments,
                    MarkAll(leftParagraphs[leftStart + offset], DiffKind.Removed));
            }

            for (var offset = pairedCount; offset < rightCount; offset++)
            {
                cancellationToken.ThrowIfCancellationRequested();
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
        List<string> rightParagraphs,
        CancellationToken cancellationToken)
    {
        const long maxMatrixCells = 250_000;
        if ((long)leftParagraphs.Count * rightParagraphs.Count > maxMatrixCells)
        {
            return [];
        }

        var leftValues = new List<string>(leftParagraphs.Count);
        foreach (var paragraph in leftParagraphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            leftValues.Add(Normalize(paragraph));
        }

        var rightValues = new List<string>(rightParagraphs.Count);
        foreach (var paragraph in rightParagraphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rightValues.Add(Normalize(paragraph));
        }

        var lengths = new int[leftValues.Count + 1, rightValues.Count + 1];

        for (var leftIndex = leftValues.Count - 1; leftIndex >= 0; leftIndex--)
        {
            cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
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
        string rightText,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var leftTokens = Tokenize(leftText);
        cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
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
        [.. DiffTokenRegex().Matches(text).Select(match => match.Value)];

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
