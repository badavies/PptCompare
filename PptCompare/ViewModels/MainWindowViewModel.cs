using System.Collections.ObjectModel;
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
    private readonly AsyncRelayCommand _openPresentationCommand;
    private readonly AsyncRelayCommand _compareCommand;
    private readonly CancellationTokenSource _lifetimeCancellation = new();
    private LoadedPresentation? _leftPresentation;
    private LoadedPresentation? _rightPresentation;
    private bool _isBusy;
    private bool _disposed;
    private VersionDescriptor? _selectedLeftVersion;
    private VersionDescriptor? _selectedRightVersion;
    private SlideComparisonItem? _selectedSlide;
    private string _statusMessage = "Ready — choose a presentation or review the sample comparison.";
    private string _leftPaneHeading = "Version 7 · 09 Sep 2026, 16:42";
    private string _rightPaneHeading = "Version 8 · Today, 09:18";
    private string _changedCountText = "3 changed";
    private string _addedCountText = "1 added";
    private string _removedCountText = "0 removed";

    public MainWindowViewModel(
        IFilePickerService filePicker,
        IPresentationSourceService presentationSource,
        IPresentationComparisonService comparisonService)
    {
        _filePicker = filePicker;
        _presentationSource = presentationSource;
        _comparisonService = comparisonService;

        LeftVersions =
        [
            new VersionDescriptor("v7", "Version 7 · 09 Sep 2026, 16:42"),
            new VersionDescriptor("v6", "Version 6 · 08 Sep 2026, 11:05")
        ];

        RightVersions =
        [
            new VersionDescriptor("v8", "Version 8 · Today, 09:18"),
            new VersionDescriptor("current", "Current local copy")
        ];

        Slides = CreateSampleSlides();
        _selectedLeftVersion = LeftVersions[0];
        _selectedRightVersion = RightVersions[0];
        _selectedSlide = Slides[0];

        _openPresentationCommand = new AsyncRelayCommand(
            OpenPresentationAsync,
            _ => !_isBusy,
            HandleUnexpectedCommandException);
        _compareCommand = new AsyncRelayCommand(
            CompareAsync,
            _ => !_isBusy && _leftPresentation is not null && _rightPresentation is not null,
            HandleUnexpectedCommandException);
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

    public ICommand OpenPresentationCommand => _openPresentationCommand;
    public ICommand CompareCommand => _compareCommand;

    private async Task OpenPresentationAsync(object? parameter)
    {
        if (!BeginOperation())
        {
            return;
        }

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

            StatusMessage = _leftPresentation is not null && _rightPresentation is not null
                ? $"Loaded {loaded.Slides.Count} slides from {fileName}. Ready to compare."
                : $"Loaded {loaded.Slides.Count} slides from {fileName}. Choose the other presentation.";
        }
        catch (PresentationLoadException exception)
        {
            StatusMessage = exception.Message;
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
            StatusMessage = "Operation cancelled.";
        }
        finally
        {
            EndOperation();
        }
    }

    private async Task CompareAsync(object? parameter)
    {
        if (!BeginOperation())
        {
            return;
        }

        try
        {
            if (_leftPresentation is null || _rightPresentation is null)
            {
                StatusMessage = "Open both presentations before comparing.";
                return;
            }

            StatusMessage = "Comparing slide text…";
            var result = await _comparisonService.CompareAsync(
                _leftPresentation,
                _rightPresentation,
                _lifetimeCancellation.Token);

            Slides.Clear();
            foreach (var slide in result.Slides)
            {
                Slides.Add(slide);
            }

            SelectedSlide = Slides.FirstOrDefault();
            ChangedCountText = $"{result.ChangedSlides} changed";
            AddedCountText = $"{result.AddedSlides} added";
            RemovedCountText = $"{result.RemovedSlides} removed";
            StatusMessage =
                $"Comparison complete — {result.ChangedSlides} changed, {result.AddedSlides} added, {result.RemovedSlides} removed.";
        }
        catch (OperationCanceledException) when (_lifetimeCancellation.IsCancellationRequested)
        {
            StatusMessage = "Comparison cancelled.";
        }
        finally
        {
            EndOperation();
        }
    }

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
    }

    private void RaiseCommandCanExecuteChanged()
    {
        _openPresentationCommand.RaiseCanExecuteChanged();
        _compareCommand.RaiseCanExecuteChanged();
    }

    private void HandleUnexpectedCommandException(Exception exception)
    {
        StatusMessage =
            "An unexpected error occurred. The operation was stopped and no presentation was modified.";
        System.Diagnostics.Debug.WriteLine(exception);
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
        GC.SuppressFinalize(this);
    }

    private static ObservableCollection<SlideComparisonItem> CreateSampleSlides() =>
    [
        new()
        {
            Number = 1,
            Title = "Executive summary",
            ChangeKind = "Changed",
            ChangeColor = "#C77B16",
            ChangeSummary = "Headline wording and forecast figure changed.",
            LeftTitle = "Q3 Business Review",
            LeftBody = "Revenue performance remained resilient across the core portfolio. The outlook assumes 8% year-on-year growth.",
            LeftCallout = "Forecast: £42.0m",
            RightTitle = "Q3 Performance Review",
            RightBody = "Revenue performance remained resilient across the core portfolio. The updated outlook assumes 11% year-on-year growth.",
            RightCallout = "Forecast: £44.5m"
        },
        new()
        {
            Number = 2,
            Title = "Market overview",
            ChangeKind = "Unchanged",
            ChangeColor = "#94A3B8",
            ChangeSummary = "No material changes detected.",
            LeftTitle = "Market overview",
            LeftBody = "Demand remains stable across priority segments, with continued strength in enterprise accounts.",
            LeftCallout = "No differences detected",
            RightTitle = "Market overview",
            RightBody = "Demand remains stable across priority segments, with continued strength in enterprise accounts.",
            RightCallout = "No differences detected"
        },
        new()
        {
            Number = 3,
            Title = "Revenue bridge",
            ChangeKind = "Changed",
            ChangeColor = "#C77B16",
            ChangeSummary = "Two values and one annotation changed.",
            LeftTitle = "Revenue bridge",
            LeftBody = "Base revenue £38.7m\nOrganic growth +£2.1m\nNew business +£1.2m",
            LeftCallout = "Net revenue: £42.0m",
            RightTitle = "Revenue bridge",
            RightBody = "Base revenue £38.7m\nOrganic growth +£3.4m\nNew business +£2.4m",
            RightCallout = "Net revenue: £44.5m"
        },
        new()
        {
            Number = 4,
            Title = "Customer highlights",
            ChangeKind = "Added",
            ChangeColor = "#15803D",
            ChangeSummary = "New slide added in the right-hand version.",
            LeftTitle = "Slide not present",
            LeftBody = "This slide does not exist in the selected left version.",
            LeftCallout = "Added on the right",
            RightTitle = "Customer highlights",
            RightBody = "Three strategic renewals completed, with improved retention across the enterprise segment.",
            RightCallout = "New content"
        },
        new()
        {
            Number = 5,
            Title = "Next steps",
            ChangeKind = "Changed",
            ChangeColor = "#C77B16",
            ChangeSummary = "Milestone date moved by two weeks.",
            LeftTitle = "Next steps",
            LeftBody = "Complete the operating review and circulate recommendations by 18 September.",
            LeftCallout = "Decision meeting: 22 September",
            RightTitle = "Next steps",
            RightBody = "Complete the operating review and circulate recommendations by 2 October.",
            RightCallout = "Decision meeting: 6 October"
        }
    ];
}
