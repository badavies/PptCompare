namespace PptCompare.Models;

public sealed record PresentationReference(string Location, string DisplayName);

public sealed record VersionDescriptor(string Id, string DisplayName, DateTimeOffset? ModifiedAt = null);

public sealed record PresentationSlide(
    int Number,
    string Title,
    IReadOnlyList<string> Paragraphs);

public sealed record LoadedPresentation(
    PresentationReference Source,
    IReadOnlyList<PresentationSlide> Slides);

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
    int RemovedSlides);

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
}
