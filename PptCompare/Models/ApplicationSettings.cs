namespace PptCompare.Models;

public sealed record PresentationReadLimits(
    long MaxFileBytes = 250L * 1024 * 1024,
    long MaxCharactersPerPart = 5_000_000,
    int MaxSlides = 2_000,
    int MaxElementsPerSlide = 10_000,
    int MaxTableCellsPerSlide = 10_000,
    int MaxParagraphsPerSlide = 10_000,
    int MaxCharactersPerSlide = 500_000,
    long MaxTotalCharacters = 20_000_000,
    long MaxRelatedPartBytes = 100L * 1024 * 1024,
    int MaxRelationshipReferences = 100_000,
    int MaxUniqueRelatedParts = 10_000,
    long MaxTotalRelatedPartBytes = 500L * 1024 * 1024,
    int MaxPreviewImageBytes = 25 * 1024 * 1024,
    long MaxTotalPreviewImageBytes = 100L * 1024 * 1024);

public sealed record ApplicationSettings(
    double OutputFontSizePoints = 10,
    int MaxFileSizeMb = 250,
    int MaxSlides = 2_000,
    int MaxElementsPerSlide = 10_000,
    int MaxTableCellsPerSlide = 10_000,
    int MaxParagraphsPerSlide = 10_000,
    int MaxCharactersPerSlide = 500_000,
    int MaxTotalCharacters = 20_000_000,
    int MaxEmbeddedItemSizeMb = 100,
    int MaxPreviewImageSizeMb = 25,
    int MaxTotalPreviewImageSizeMb = 100,
    bool UsePowerPointRendering = true,
    bool EnableDebugLogging = false)
{
    public const double MinOutputFontSizePoints = 8;
    public const double MaxOutputFontSizePoints = 24;
    public const int MaxAllowedFileSizeMb = 1_000;
    public const int MaxAllowedSlides = 5_000;
    public const int MaxAllowedElementsPerSlide = 50_000;
    public const int MaxAllowedTableCellsPerSlide = 50_000;
    public const int MaxAllowedParagraphsPerSlide = 50_000;
    public const int MaxAllowedCharactersPerSlide = 2_000_000;
    public const int MaxAllowedTotalCharacters = 100_000_000;
    public const int MaxAllowedEmbeddedItemSizeMb = 500;
    public const int MaxAllowedPreviewImageSizeMb = 100;
    public const int MaxAllowedTotalPreviewImageSizeMb = 500;

    public PresentationReadLimits ToReadLimits() => new(
        MaxFileBytes: MegabytesToBytes(MaxFileSizeMb),
        MaxCharactersPerPart: Math.Max(5_000_000, MaxCharactersPerSlide),
        MaxSlides: MaxSlides,
        MaxElementsPerSlide: MaxElementsPerSlide,
        MaxTableCellsPerSlide: MaxTableCellsPerSlide,
        MaxParagraphsPerSlide: MaxParagraphsPerSlide,
        MaxCharactersPerSlide: MaxCharactersPerSlide,
        MaxTotalCharacters: MaxTotalCharacters,
        MaxRelatedPartBytes: MegabytesToBytes(MaxEmbeddedItemSizeMb),
        MaxPreviewImageBytes: checked(MaxPreviewImageSizeMb * 1024 * 1024),
        MaxTotalPreviewImageBytes: MegabytesToBytes(MaxTotalPreviewImageSizeMb));

    public static bool TryValidate(ApplicationSettings? settings, out string error)
    {
        if (settings is null)
        {
            error = "The settings file did not contain any settings.";
            return false;
        }

        if (!double.IsFinite(settings.OutputFontSizePoints) ||
            settings.OutputFontSizePoints is < MinOutputFontSizePoints or > MaxOutputFontSizePoints)
        {
            error = $"Text output size must be between {MinOutputFontSizePoints:N0} and {MaxOutputFontSizePoints:N0} points.";
            return false;
        }

        if (!InRange(settings.MaxFileSizeMb, 1, MaxAllowedFileSizeMb))
        {
            error = $"Maximum presentation size must be between 1 and {MaxAllowedFileSizeMb:N0} MB.";
            return false;
        }

        if (!InRange(settings.MaxSlides, 1, MaxAllowedSlides))
        {
            error = $"Maximum slides must be between 1 and {MaxAllowedSlides:N0}.";
            return false;
        }

        if (!InRange(settings.MaxElementsPerSlide, 1, MaxAllowedElementsPerSlide))
        {
            error = $"Maximum elements per slide must be between 1 and {MaxAllowedElementsPerSlide:N0}.";
            return false;
        }

        if (!InRange(settings.MaxTableCellsPerSlide, 1, MaxAllowedTableCellsPerSlide))
        {
            error = $"Maximum table cells per slide must be between 1 and {MaxAllowedTableCellsPerSlide:N0}.";
            return false;
        }

        if (!InRange(settings.MaxParagraphsPerSlide, 1, MaxAllowedParagraphsPerSlide))
        {
            error = $"Maximum paragraphs per slide must be between 1 and {MaxAllowedParagraphsPerSlide:N0}.";
            return false;
        }

        if (!InRange(settings.MaxCharactersPerSlide, 1_000, MaxAllowedCharactersPerSlide))
        {
            error = $"Maximum characters per slide must be between 1,000 and {MaxAllowedCharactersPerSlide:N0}.";
            return false;
        }

        if (!InRange(settings.MaxTotalCharacters, settings.MaxCharactersPerSlide, MaxAllowedTotalCharacters))
        {
            error = $"Maximum total characters must be at least the per-slide limit and no more than {MaxAllowedTotalCharacters:N0}.";
            return false;
        }

        if (!InRange(settings.MaxEmbeddedItemSizeMb, 1, MaxAllowedEmbeddedItemSizeMb))
        {
            error = $"Maximum embedded item size must be between 1 and {MaxAllowedEmbeddedItemSizeMb:N0} MB.";
            return false;
        }

        if (!InRange(settings.MaxPreviewImageSizeMb, 1, MaxAllowedPreviewImageSizeMb))
        {
            error = $"Maximum preview image size must be between 1 and {MaxAllowedPreviewImageSizeMb:N0} MB.";
            return false;
        }

        if (!InRange(
                settings.MaxTotalPreviewImageSizeMb,
                settings.MaxPreviewImageSizeMb,
                MaxAllowedTotalPreviewImageSizeMb))
        {
            error = $"Maximum total preview image size must be at least the per-image limit and no more than {MaxAllowedTotalPreviewImageSizeMb:N0} MB.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    private static bool InRange(int value, int minimum, int maximum) =>
        value >= minimum && value <= maximum;

    private static long MegabytesToBytes(int megabytes) =>
        checked((long)megabytes * 1024 * 1024);
}
