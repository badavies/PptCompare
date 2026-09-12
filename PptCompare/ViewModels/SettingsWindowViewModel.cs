using System.Globalization;
using PptCompare.Infrastructure;
using PptCompare.Models;

namespace PptCompare.ViewModels;

public sealed class SettingsWindowViewModel : ObservableObject
{
    public SettingsWindowViewModel(ApplicationSettings settings) => Load(settings);

    public string OutputFontSizePoints { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxFileSizeMb { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxSlides { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxElementsPerSlide { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxTableCellsPerSlide { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxParagraphsPerSlide { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxCharactersPerSlide { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxTotalCharacters { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxEmbeddedItemSizeMb { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxPreviewImageSizeMb { get; set => SetProperty(ref field, value); } = string.Empty;

    public string MaxTotalPreviewImageSizeMb { get; set => SetProperty(ref field, value); } = string.Empty;

    public bool UsePowerPointRendering
    {
        get;
        set => SetProperty(ref field, value);
    }

    public bool EnableDebugLogging
    {
        get;
        set => SetProperty(ref field, value);
    }

    public void RestoreDefaults() => Load(new ApplicationSettings());

    public bool TryCreateSettings(out ApplicationSettings? settings, out string error)
    {
        settings = null;
        if (!TryReadDouble(OutputFontSizePoints, "Text output size", out var outputFontSize, out error) ||
            !TryReadInt(MaxFileSizeMb, "Maximum presentation size", out var maxFileSize, out error) ||
            !TryReadInt(MaxSlides, "Maximum slides", out var maxSlides, out error) ||
            !TryReadInt(MaxElementsPerSlide, "Maximum elements per slide", out var maxElements, out error) ||
            !TryReadInt(MaxTableCellsPerSlide, "Maximum table cells per slide", out var maxTableCells, out error) ||
            !TryReadInt(MaxParagraphsPerSlide, "Maximum paragraphs per slide", out var maxParagraphs, out error) ||
            !TryReadInt(MaxCharactersPerSlide, "Maximum characters per slide", out var maxCharacters, out error) ||
            !TryReadInt(MaxTotalCharacters, "Maximum total characters", out var maxTotalCharacters, out error) ||
            !TryReadInt(MaxEmbeddedItemSizeMb, "Maximum embedded item size", out var maxEmbeddedItemSize, out error) ||
            !TryReadInt(MaxPreviewImageSizeMb, "Maximum preview image size", out var maxPreviewImageSize, out error) ||
            !TryReadInt(MaxTotalPreviewImageSizeMb, "Maximum total preview size", out var maxTotalPreviewImageSize, out error))
        {
            return false;
        }

        var candidate = new ApplicationSettings(
            outputFontSize,
            maxFileSize,
            maxSlides,
            maxElements,
            maxTableCells,
            maxParagraphs,
            maxCharacters,
            maxTotalCharacters,
            maxEmbeddedItemSize,
            maxPreviewImageSize,
            maxTotalPreviewImageSize,
            UsePowerPointRendering,
            EnableDebugLogging);
        if (!ApplicationSettings.TryValidate(candidate, out error))
        {
            return false;
        }

        settings = candidate;
        return true;
    }

    private void Load(ApplicationSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        OutputFontSizePoints = settings.OutputFontSizePoints.ToString("0.##", CultureInfo.CurrentCulture);
        MaxFileSizeMb = Format(settings.MaxFileSizeMb);
        MaxSlides = Format(settings.MaxSlides);
        MaxElementsPerSlide = Format(settings.MaxElementsPerSlide);
        MaxTableCellsPerSlide = Format(settings.MaxTableCellsPerSlide);
        MaxParagraphsPerSlide = Format(settings.MaxParagraphsPerSlide);
        MaxCharactersPerSlide = Format(settings.MaxCharactersPerSlide);
        MaxTotalCharacters = Format(settings.MaxTotalCharacters);
        MaxEmbeddedItemSizeMb = Format(settings.MaxEmbeddedItemSizeMb);
        MaxPreviewImageSizeMb = Format(settings.MaxPreviewImageSizeMb);
        MaxTotalPreviewImageSizeMb = Format(settings.MaxTotalPreviewImageSizeMb);
        UsePowerPointRendering = settings.UsePowerPointRendering;
        EnableDebugLogging = settings.EnableDebugLogging;
    }

    private static string Format(int value) => value.ToString(CultureInfo.CurrentCulture);

    private static bool TryReadInt(string text, string fieldName, out int value, out string error)
    {
        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out value))
        {
            error = $"{fieldName} must be a whole number.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    private static bool TryReadDouble(string text, string fieldName, out double value, out string error)
    {
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
            !double.IsFinite(value))
        {
            error = $"{fieldName} must be a valid number.";
            return false;
        }

        error = string.Empty;
        return true;
    }
}
