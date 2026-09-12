using System.Globalization;
using PptCompare.Infrastructure;
using PptCompare.Models;

namespace PptCompare.ViewModels;

public sealed class SettingsWindowViewModel : ObservableObject
{
    private string _outputFontSizePoints = string.Empty;
    private string _maxFileSizeMb = string.Empty;
    private string _maxSlides = string.Empty;
    private string _maxElementsPerSlide = string.Empty;
    private string _maxTableCellsPerSlide = string.Empty;
    private string _maxParagraphsPerSlide = string.Empty;
    private string _maxCharactersPerSlide = string.Empty;
    private string _maxTotalCharacters = string.Empty;
    private string _maxEmbeddedItemSizeMb = string.Empty;
    private string _maxPreviewImageSizeMb = string.Empty;
    private string _maxTotalPreviewImageSizeMb = string.Empty;
    private bool _usePowerPointRendering;
    private bool _enableDebugLogging;

    public SettingsWindowViewModel(ApplicationSettings settings) => Load(settings);

    public string OutputFontSizePoints
    {
        get => _outputFontSizePoints;
        set => SetProperty(ref _outputFontSizePoints, value);
    }

    public string MaxFileSizeMb
    {
        get => _maxFileSizeMb;
        set => SetProperty(ref _maxFileSizeMb, value);
    }

    public string MaxSlides
    {
        get => _maxSlides;
        set => SetProperty(ref _maxSlides, value);
    }

    public string MaxElementsPerSlide
    {
        get => _maxElementsPerSlide;
        set => SetProperty(ref _maxElementsPerSlide, value);
    }

    public string MaxTableCellsPerSlide
    {
        get => _maxTableCellsPerSlide;
        set => SetProperty(ref _maxTableCellsPerSlide, value);
    }

    public string MaxParagraphsPerSlide
    {
        get => _maxParagraphsPerSlide;
        set => SetProperty(ref _maxParagraphsPerSlide, value);
    }

    public string MaxCharactersPerSlide
    {
        get => _maxCharactersPerSlide;
        set => SetProperty(ref _maxCharactersPerSlide, value);
    }

    public string MaxTotalCharacters
    {
        get => _maxTotalCharacters;
        set => SetProperty(ref _maxTotalCharacters, value);
    }

    public string MaxEmbeddedItemSizeMb
    {
        get => _maxEmbeddedItemSizeMb;
        set => SetProperty(ref _maxEmbeddedItemSizeMb, value);
    }

    public string MaxPreviewImageSizeMb
    {
        get => _maxPreviewImageSizeMb;
        set => SetProperty(ref _maxPreviewImageSizeMb, value);
    }

    public string MaxTotalPreviewImageSizeMb
    {
        get => _maxTotalPreviewImageSizeMb;
        set => SetProperty(ref _maxTotalPreviewImageSizeMb, value);
    }

    public bool UsePowerPointRendering
    {
        get => _usePowerPointRendering;
        set => SetProperty(ref _usePowerPointRendering, value);
    }

    public bool EnableDebugLogging
    {
        get => _enableDebugLogging;
        set => SetProperty(ref _enableDebugLogging, value);
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
