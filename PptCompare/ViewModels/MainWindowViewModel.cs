using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Windows.Input;
using PptCompare.Infrastructure;
using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare.ViewModels;

public sealed class MainWindowViewModel : ObservableObject, IDisposable
{
    private readonly IFilePickerService _filePicker;
    private readonly IPresentationSourceService _presentationSource;
    private readonly IPresentationComparisonService _comparisonService;
    private readonly IPresentationRenderer _renderer;
    private readonly IApplicationDiagnostics _diagnostics;
    private readonly IApplicationSettingsService _settingsService;
    private readonly ISettingsDialogService _settingsDialog;
    private readonly IAboutDialogService _aboutDialog;
    private readonly AsyncRelayCommand _openPresentationCommand;
    private readonly AsyncRelayCommand _compareCommand;
    private readonly AsyncRelayCommand _swapSidesCommand;
    private readonly RelayCommand _settingsCommand;
    private readonly RelayCommand _aboutCommand;
    private readonly CancellationTokenSource _lifetimeCancellation = new();
    private readonly object _previewRenderGate = new();
    private readonly HashSet<LoadedPresentation> _presentationsBeingRendered =
        new(ReferenceEqualityComparer.Instance);
    private LoadedPresentation? _leftPresentation;
    private LoadedPresentation? _rightPresentation;
    private bool _isBusy;
    private bool _disposed;
    private bool _hasCompared;
    private bool _comparisonRefreshPending;
    private VersionDescriptor? _selectedLeftVersion;
    private VersionDescriptor? _selectedRightVersion;
    private SlideComparisonItem? _selectedSlide;
    private string _statusMessage = "Ready — select a presentation on each side.";
    private string _leftPaneHeading = "No left presentation selected";
    private string _rightPaneHeading = "No right presentation selected";
    private string _changedCountText = "0 changed";
    private string _addedCountText = "0 added";
    private string _removedCountText = "0 removed";
    private string _movedCountText = "0 moved";
    private ApplicationSettings _settings;
    private double _outputTextFontSize;

    public MainWindowViewModel(
        IFilePickerService filePicker,
        IPresentationSourceService presentationSource,
        IPresentationComparisonService comparisonService,
        IPresentationRenderer renderer,
        IApplicationDiagnostics diagnostics,
        IApplicationSettingsService settingsService,
        ISettingsDialogService settingsDialog,
        IAboutDialogService aboutDialog,
        ApplicationSettings settings)
    {
        ArgumentNullException.ThrowIfNull(filePicker);
        ArgumentNullException.ThrowIfNull(presentationSource);
        ArgumentNullException.ThrowIfNull(comparisonService);
        ArgumentNullException.ThrowIfNull(renderer);
        ArgumentNullException.ThrowIfNull(diagnostics);
        ArgumentNullException.ThrowIfNull(settingsService);
        ArgumentNullException.ThrowIfNull(settingsDialog);
        ArgumentNullException.ThrowIfNull(aboutDialog);
        ArgumentNullException.ThrowIfNull(settings);
        _filePicker = filePicker;
        _presentationSource = presentationSource;
        _comparisonService = comparisonService;
        _renderer = renderer;
        _diagnostics = diagnostics;
        _settingsService = settingsService;
        _settingsDialog = settingsDialog;
        _aboutDialog = aboutDialog;
        _settings = settings;
        _outputTextFontSize = PointsToDeviceIndependentPixels(settings.OutputFontSizePoints);

        LeftVersions = [];
        RightVersions = [];
        Slides = [];

        _openPresentationCommand = new AsyncRelayCommand(
            OpenPresentationAsync,
            _ => !_isBusy,
            HandleUnexpectedCommandException);
        _compareCommand = new AsyncRelayCommand(
            CompareAsync,
            _ => !_isBusy && _leftPresentation is not null && _rightPresentation is not null,
            HandleUnexpectedCommandException);
        _swapSidesCommand = new AsyncRelayCommand(
            SwapSidesAsync,
            _ => !_isBusy,
            HandleUnexpectedCommandException);
        _settingsCommand = new RelayCommand(OpenSettings, _ => !_isBusy);
        _aboutCommand = new RelayCommand(OpenAbout, _ => !_isBusy);
    }

    public ObservableCollection<VersionDescriptor> LeftVersions { get; }
    public ObservableCollection<VersionDescriptor> RightVersions { get; }
    public ObservableCollection<SlideComparisonItem> Slides { get; }

    public VersionDescriptor? SelectedLeftVersion
    {
        get => _selectedLeftVersion;
        set
        {
            if (SetProperty(ref _selectedLeftVersion, value) && value is not null)
            {
                LeftPaneHeading = value.DisplayName;
            }
        }
    }

    public VersionDescriptor? SelectedRightVersion
    {
        get => _selectedRightVersion;
        set
        {
            if (SetProperty(ref _selectedRightVersion, value) && value is not null)
            {
                RightPaneHeading = value.DisplayName;
            }
        }
    }

    public SlideComparisonItem? SelectedSlide
    {
        get => _selectedSlide;
        set => SetProperty(ref _selectedSlide, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => SetProperty(ref _statusMessage, value);
    }

    public string LeftPaneHeading
    {
        get => _leftPaneHeading;
        private set => SetProperty(ref _leftPaneHeading, value);
    }

    public string RightPaneHeading
    {
        get => _rightPaneHeading;
        private set => SetProperty(ref _rightPaneHeading, value);
    }

    public string ChangedCountText
    {
        get => _changedCountText;
        private set => SetProperty(ref _changedCountText, value);
    }

    public string AddedCountText
    {
        get => _addedCountText;
        private set => SetProperty(ref _addedCountText, value);
    }

    public string RemovedCountText
    {
        get => _removedCountText;
        private set => SetProperty(ref _removedCountText, value);
    }

    public string MovedCountText
    {
        get => _movedCountText;
        private set => SetProperty(ref _movedCountText, value);
    }

    public double OutputTextFontSize
    {
        get => _outputTextFontSize;
        private set => SetProperty(ref _outputTextFontSize, value);
    }

    public ICommand OpenPresentationCommand => _openPresentationCommand;
    public ICommand CompareCommand => _compareCommand;
    public ICommand SwapSidesCommand => _swapSidesCommand;
    public ICommand SettingsCommand => _settingsCommand;
    public ICommand AboutCommand => _aboutCommand;

    private async Task OpenPresentationAsync(object? parameter)
    {
        if (!BeginOperation())
        {
            return;
        }

        LoadedPresentation? previewToStart = null;
        try
        {
            var path = _filePicker.PickPresentation();
            if (path is null)
            {
                StatusMessage = "Open cancelled.";
                return;
            }

            var side = parameter as string;
            if (side is null)
            {
                side = _leftPresentation is null ? "Left" : "Right";
            }

            var fileName = Path.GetFileName(path);
            var displayName = $"Local · {fileName}";
            var version = new VersionDescriptor($"local:{path}", displayName, DateTimeOffset.Now);
            StatusMessage = $"Reading {fileName}…";
            var loaded = await _presentationSource.LoadAsync(
                new PresentationReference(path, fileName),
                _lifetimeCancellation.Token);
            _diagnostics.RecordEvent("PresentationLoaded", $"slides={loaded.Slides.Count}");

            if (string.Equals(side, "Left", StringComparison.OrdinalIgnoreCase))
            {
                _leftPresentation = loaded;
                LeftVersions.Clear();
                LeftVersions.Add(version);
                SelectedLeftVersion = version;
            }
            else
            {
                _rightPresentation = loaded;
                RightVersions.Clear();
                RightVersions.Add(version);
                SelectedRightVersion = version;
            }

            _hasCompared = false;
            _comparisonRefreshPending = false;
            StatusMessage = BuildLoadedStatus(loaded, fileName);
            if (_settings.UsePowerPointRendering)
            {
                previewToStart = loaded;
            }
        }
        catch (PresentationLoadException exception)
        {
            _diagnostics.RecordException("PresentationLoadRejected", exception);
            StatusMessage = exception.Message;
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
            StatusMessage = "Operation cancelled.";
        }
        finally
        {
            EndOperation();
            if (previewToStart is not null)
            {
                QueuePreviewRendering(previewToStart);
            }
        }
    }

    private async Task CompareAsync(object? parameter)
    {
        if (!BeginOperation())
        {
            return;
        }

        var retryMissingPreviews = false;
        try
        {
            if (_leftPresentation is null || _rightPresentation is null)
            {
                StatusMessage = "Open both presentations before comparing.";
                return;
            }

            StatusMessage = "Comparing presentations…";
            var result = await _comparisonService.CompareAsync(
                _leftPresentation,
                _rightPresentation,
                _lifetimeCancellation.Token);

            ApplyComparisonResult(result);
            _hasCompared = true;
            _diagnostics.RecordEvent(
                "ComparisonCompleted",
                $"changed={result.ChangedSlides};moved={result.MovedSlides};added={result.AddedSlides};removed={result.RemovedSlides}");
            StatusMessage =
                $"Comparison complete — {result.ChangedSlides} changed, {result.MovedSlides} moved, {result.AddedSlides} added, {result.RemovedSlides} removed.";
            retryMissingPreviews = _settings.UsePowerPointRendering;
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
            StatusMessage = "Comparison cancelled.";
        }
        finally
        {
            EndOperation();
            if (retryMissingPreviews)
            {
                StartMissingPreviewRendering();
            }
        }
    }

    private async Task SwapSidesAsync(object? parameter)
    {
        if (!BeginOperation())
        {
            return;
        }

        try
        {
            var selectedSlideNumber = SelectedSlide?.Number;
            var previousLeftHeading = LeftPaneHeading;
            var previousRightHeading = RightPaneHeading;
            var previousLeftSelected = SelectedLeftVersion;
            var previousRightSelected = SelectedRightVersion;
            var previousLeftVersions = LeftVersions.ToList();
            var previousRightVersions = RightVersions.ToList();

            (_leftPresentation, _rightPresentation) = (_rightPresentation, _leftPresentation);

            SelectedLeftVersion = null;
            SelectedRightVersion = null;
            LeftVersions.Clear();
            foreach (var version in previousRightVersions)
            {
                LeftVersions.Add(version);
            }

            RightVersions.Clear();
            foreach (var version in previousLeftVersions)
            {
                RightVersions.Add(version);
            }

            SelectedLeftVersion = previousRightSelected ?? LeftVersions.FirstOrDefault();
            SelectedRightVersion = previousLeftSelected ?? RightVersions.FirstOrDefault();
            if (SelectedLeftVersion is null)
            {
                LeftPaneHeading = previousRightHeading;
            }

            if (SelectedRightVersion is null)
            {
                RightPaneHeading = previousLeftHeading;
            }

            if (_hasCompared && _leftPresentation is not null && _rightPresentation is not null)
            {
                StatusMessage = "Switching comparison sides…";
                var result = await _comparisonService.CompareAsync(
                    _leftPresentation,
                    _rightPresentation,
                    _lifetimeCancellation.Token);
                ApplyComparisonResult(result, selectedSlideNumber);
                StatusMessage = "Left and right presentations switched.";
            }
            else
            {
                SwapDisplayedComparison(selectedSlideNumber);
                StatusMessage = "Left and right presentations switched.";
            }
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
            StatusMessage = "Switch cancelled.";
        }
        finally
        {
            EndOperation();
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "A UI command boundary must convert unexpected settings errors into a safe status message.")]
    private void OpenSettings(object? parameter)
    {
        try
        {
            var previousSettings = _settings;
            var updatedSettings = _settingsDialog.EditSettings(previousSettings);
            if (updatedSettings is null)
            {
                return;
            }

            _settingsService.Save(updatedSettings);
            if (_diagnostics is IDetailedApplicationDiagnostics detailedDiagnostics)
            {
                detailedDiagnostics.SetDebugLoggingEnabled(updatedSettings.EnableDebugLogging);
            }

            if (_presentationSource is IConfigurablePresentationSourceService configurableSource)
            {
                configurableSource.UpdateLimits(updatedSettings.ToReadLimits());
            }

            _settings = updatedSettings;
            OutputTextFontSize = PointsToDeviceIndependentPixels(updatedSettings.OutputFontSizePoints);
            StatusMessage = "Settings saved. File safety limits apply when the next presentation is opened.";

            if (!previousSettings.UsePowerPointRendering && updatedSettings.UsePowerPointRendering)
            {
                StartMissingPreviewRendering();
            }
        }
        catch (SettingsPersistenceException exception)
        {
            Debug.WriteLine(exception);
            _diagnostics.RecordException("SettingsCommandFailed", exception);
            StatusMessage = exception.Message;
        }
        catch (Exception exception)
        {
            HandleUnexpectedCommandException(exception);
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "A UI command boundary must convert unexpected dialog errors into a safe status message.")]
    private void OpenAbout(object? parameter)
    {
        try
        {
            _aboutDialog.ShowAbout();
        }
        catch (Exception exception)
        {
            HandleUnexpectedCommandException(exception);
        }
    }

    private void StartMissingPreviewRendering()
    {
        if (_leftPresentation is { } left && left.Slides.Any(slide => slide.RenderedImagePath is null))
        {
            QueuePreviewRendering(left);
        }

        if (_rightPresentation is { } right &&
            !ReferenceEquals(right, _leftPresentation) &&
            right.Slides.Any(slide => slide.RenderedImagePath is null))
        {
            QueuePreviewRendering(right);
        }
    }

    private void QueuePreviewRendering(LoadedPresentation presentation)
    {
        lock (_previewRenderGate)
        {
            if (!_presentationsBeingRendered.Add(presentation))
            {
                return;
            }
        }

        _ = RenderPresentationInBackgroundAsync(presentation);
    }

    private string BuildLoadedStatus(LoadedPresentation loaded, string fileName)
    {
        var nextAction = _leftPresentation is not null && _rightPresentation is not null
            ? "Ready to compare."
            : "Choose the other presentation.";
        var previewStatus = _settings.UsePowerPointRendering
            ? " Preparing the PowerPoint preview in the background."
            : " The built-in preview is ready.";
        return $"Loaded {loaded.Slides.Count} slides from {fileName}. {nextAction}{previewStatus}";
    }

    private static double PointsToDeviceIndependentPixels(double points) => points * 96d / 72d;

    private void SwapDisplayedComparison(int? selectedSlideNumber)
    {
        var swappedSlides = Slides.Select(SwapComparisonItem).ToList();
        Slides.Clear();
        foreach (var slide in swappedSlides)
        {
            Slides.Add(slide);
        }

        SelectedSlide = selectedSlideNumber is null
            ? Slides.FirstOrDefault()
            : Slides.FirstOrDefault(slide => slide.Number == selectedSlideNumber) ?? Slides.FirstOrDefault();
    }

    private static SlideComparisonItem SwapComparisonItem(SlideComparisonItem item)
    {
        var reversedChangeKind = item.ChangeKind switch
        {
            "Added" => "Removed",
            "Removed" => "Added",
            _ => item.ChangeKind
        };

        return new SlideComparisonItem
        {
            Number = item.Number,
            Title = item.LeftSlide?.Title ?? item.LeftTitle,
            ChangeKind = reversedChangeKind,
            ChangeColor = reversedChangeKind switch
            {
                "Added" => "#15803D",
                "Removed" => "#C2413B",
                "Changed" => "#C77B16",
                _ => "#94A3B8"
            },
            ChangeSummary = item.ChangeKind switch
            {
                "Added" => "Slide exists only in the left-hand version.",
                "Removed" => "Slide exists only in the right-hand version.",
                _ => item.ChangeSummary
            },
            LeftTitle = item.RightTitle,
            LeftBody = item.RightBody,
            LeftCallout = ReverseCallout(item.RightCallout),
            RightTitle = item.LeftTitle,
            RightBody = item.LeftBody,
            RightCallout = ReverseCallout(item.LeftCallout),
            LeftTitleSegments = ReverseSegments(item.RightTitleSegments),
            LeftBodySegments = ReverseSegments(item.RightBodySegments),
            RightTitleSegments = ReverseSegments(item.LeftTitleSegments),
            RightBodySegments = ReverseSegments(item.LeftBodySegments),
            LeftTitleContent = item.RightTitleContent,
            LeftBodyContent = item.RightBodyContent,
            RightTitleContent = item.LeftTitleContent,
            RightBodyContent = item.LeftBodyContent,
            LeftSlide = item.RightSlide,
            RightSlide = item.LeftSlide,
            ElementChanges = item.ElementChanges.Select(ReverseElementChange).ToList()
        };
    }

    private static List<DiffSegment>? ReverseSegments(
        IReadOnlyList<DiffSegment>? segments) =>
        segments?.Select(segment => segment with
        {
            Kind = segment.Kind switch
            {
                DiffKind.Added => DiffKind.Removed,
                DiffKind.Removed => DiffKind.Added,
                _ => DiffKind.Unchanged
            }
        }).ToList();

    private static SlideElementChange ReverseElementChange(SlideElementChange change) =>
        change with
        {
            Kind = change.Kind switch
            {
                SlideElementChangeKind.Added => SlideElementChangeKind.Removed,
                SlideElementChangeKind.Removed => SlideElementChangeKind.Added,
                _ => change.Kind
            },
            Description = change.Kind switch
            {
                SlideElementChangeKind.Added => change.Description.Replace(
                    " added",
                    " removed",
                    StringComparison.OrdinalIgnoreCase),
                SlideElementChangeKind.Removed => change.Description.Replace(
                    " removed",
                    " added",
                    StringComparison.OrdinalIgnoreCase),
                _ => change.Description
            },
            LeftBounds = change.RightBounds,
            RightBounds = change.LeftBounds
        };

    private static string ReverseCallout(string value) =>
        value
            .Replace("Added or changed", "Swapped or changed", StringComparison.OrdinalIgnoreCase)
            .Replace("Removed or changed", "Added or changed", StringComparison.OrdinalIgnoreCase)
            .Replace("Swapped or changed", "Removed or changed", StringComparison.OrdinalIgnoreCase);

    private bool BeginOperation()
    {
        if (_isBusy)
        {
            StatusMessage = "Another operation is already in progress.";
            return false;
        }

        _isBusy = true;
        RaiseCommandCanExecuteChanged();
        return true;
    }

    private void EndOperation()
    {
        _isBusy = false;
        RaiseCommandCanExecuteChanged();
        if (_comparisonRefreshPending && _hasCompared && !_disposed)
        {
            _comparisonRefreshPending = false;
            _ = RefreshComparisonAfterRenderingAsync();
        }
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "An optional background preview must never terminate the application or become an unobserved task exception.")]
    private async Task RenderPresentationInBackgroundAsync(LoadedPresentation loaded)
    {
        try
        {
            var rendering = await _renderer.RenderAsync(
                loaded.Source.Location,
                loaded.Slides.Count,
                _lifetimeCancellation.Token);
            if (_disposed)
            {
                return;
            }

            var receivedNewPreview = rendering.SlideImages.Count > 0;
            var renderedSlides = loaded.Slides
                .Select(slide => rendering.SlideImages.TryGetValue(slide.Number, out var imagePath)
                    ? slide with { RenderedImagePath = imagePath }
                    : slide)
                .ToList();
            var renderedPresentation = loaded with
            {
                Slides = renderedSlides,
                RenderingStatus = rendering.Status
            };

            var isLeft = ReferenceEquals(_leftPresentation, loaded);
            var isRight = ReferenceEquals(_rightPresentation, loaded);
            if (!isLeft && !isRight)
            {
                return;
            }

            if (isLeft)
            {
                _leftPresentation = renderedPresentation;
            }
            else
            {
                _rightPresentation = renderedPresentation;
            }

            if (receivedNewPreview && _isBusy)
            {
                _comparisonRefreshPending = true;
            }
            else if (receivedNewPreview && _hasCompared)
            {
                await RefreshComparisonAfterRenderingAsync();
            }

            if (!_isBusy)
            {
                StatusMessage = rendering.Status;
            }
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
        }
        catch (ObjectDisposedException) when (_disposed)
        {
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            _diagnostics.RecordException("BackgroundPreviewFailed", exception);
            if (!_disposed && !_isBusy)
            {
                StatusMessage = "The PowerPoint preview was unavailable; the built-in preview remains ready.";
            }
        }
        finally
        {
            lock (_previewRenderGate)
            {
                _presentationsBeingRendered.Remove(loaded);
            }
        }
    }

    private async Task RefreshComparisonAfterRenderingAsync()
    {
        if (_disposed || _isBusy || _leftPresentation is null || _rightPresentation is null)
        {
            return;
        }

        try
        {
            var selectedNumber = SelectedSlide?.Number;
            var result = await _comparisonService.CompareAsync(
                _leftPresentation,
                _rightPresentation,
                _lifetimeCancellation.Token);
            if (_disposed)
            {
                return;
            }

            ApplyComparisonResult(result, selectedNumber);
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
        }
    }

    private void ApplyComparisonResult(
        PresentationComparisonResult result,
        int? selectedSlideNumber = null)
    {
        Slides.Clear();
        foreach (var slide in result.Slides)
        {
            Slides.Add(slide);
        }

        SelectedSlide = selectedSlideNumber is null
            ? Slides.FirstOrDefault()
            : Slides.FirstOrDefault(slide => slide.Number == selectedSlideNumber) ?? Slides.FirstOrDefault();
        ChangedCountText = $"{result.ChangedSlides} changed";
        AddedCountText = $"{result.AddedSlides} added";
        RemovedCountText = $"{result.RemovedSlides} removed";
        MovedCountText = $"{result.MovedSlides} moved";
    }

    private void RaiseCommandCanExecuteChanged()
    {
        _openPresentationCommand.RaiseCanExecuteChanged();
        _compareCommand.RaiseCanExecuteChanged();
        _swapSidesCommand.RaiseCanExecuteChanged();
        _settingsCommand.RaiseCanExecuteChanged();
        _aboutCommand.RaiseCanExecuteChanged();
    }

    private void HandleUnexpectedCommandException(Exception exception)
    {
        StatusMessage =
            "An unexpected error occurred. The operation was stopped and no presentation was modified.";
        _diagnostics.RecordException("CommandFailed", exception);
        Debug.WriteLine(exception);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _lifetimeCancellation.Cancel();
        _lifetimeCancellation.Dispose();
        _renderer.Dispose();

        GC.SuppressFinalize(this);
    }

}
