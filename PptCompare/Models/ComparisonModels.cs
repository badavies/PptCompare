namespace PptCompare.Models;

public sealed record PresentationReference(string Location, string DisplayName);

public sealed record VersionDescriptor(string Id, string DisplayName, DateTimeOffset? ModifiedAt = null);

public sealed record PresentationSlide(
    int Number,
    string Title,
    IReadOnlyList<string> Paragraphs,
    long Width,
    long Height,
    IReadOnlyList<SlideElement> Elements,
    string? RenderedImagePath = null);

public sealed record LoadedPresentation(
    PresentationReference Source,
    IReadOnlyList<PresentationSlide> Slides,
    string RenderingStatus = "Using the built-in slide preview.");

public enum SlideElementKind
{
    Text,
    Image,
    Shape,
    Connector,
    Group,
    Table,
    Chart,
    Other
}

public sealed record SlideBounds(long X, long Y, long Width, long Height);

public sealed record SlideElement(
    string Id,
    string Name,
    SlideElementKind Kind,
    SlideBounds Bounds,
    string ContentHash,
    string VisualHash,
    string Text,
    byte[]? ImageBytes = null);

public enum SlideElementChangeKind
{
    Added,
    Removed,
    Modified,
    Moved,
    Resized,
    Replaced
}

public sealed record SlideElementChange(
    SlideElementChangeKind Kind,
    SlideElementKind ElementKind,
    string Description,
    SlideBounds? LeftBounds,
    SlideBounds? RightBounds);

public enum DiffKind
{
    Unchanged,
    Added,
    Removed
}

public sealed record DiffSegment(string Text, DiffKind Kind);

public sealed record PresentationComparisonResult(
    IReadOnlyList<SlideComparisonItem> Slides,
    int ChangedSlides,
    int AddedSlides,
    int RemovedSlides,
    int MovedSlides = 0);

public sealed class SlideComparisonItem
{
    public required int Number { get; init; }
    public required string Title { get; init; }
    public required string ChangeKind { get; init; }
    public required string ChangeColor { get; init; }
    public required string ChangeSummary { get; init; }
    public required string LeftTitle { get; init; }
    public required string LeftBody { get; init; }
    public required string LeftCallout { get; init; }
    public required string RightTitle { get; init; }
    public required string RightBody { get; init; }
    public required string RightCallout { get; init; }
    public IReadOnlyList<DiffSegment>? LeftTitleSegments { get; init; }
    public IReadOnlyList<DiffSegment>? LeftBodySegments { get; init; }
    public IReadOnlyList<DiffSegment>? RightTitleSegments { get; init; }
    public IReadOnlyList<DiffSegment>? RightBodySegments { get; init; }
    public PresentationSlide? LeftSlide { get; init; }
    public PresentationSlide? RightSlide { get; init; }
    public IReadOnlyList<SlideElementChange> ElementChanges { get; init; } = [];
}
