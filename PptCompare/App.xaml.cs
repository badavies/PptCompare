using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Windows;
using System.Windows.Threading;
using PptCompare.Models;
using PptCompare.Services;
using PptCompare.ViewModels;

namespace PptCompare;

public partial class App : Application
{
    private IApplicationDiagnostics _diagnostics = NullApplicationDiagnostics.Instance;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _diagnostics = CreateDiagnostics();
        _diagnostics.RecordEvent("ApplicationStarted", $"version={ApplicationInfo.Version}");
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        IFilePickerService filePicker = new WpfFilePickerService();
        IPresentationRenderer renderer = new PowerPointPresentationRenderer(_diagnostics);
        var settingsService = new JsonApplicationSettingsService(_diagnostics);
        var settings = settingsService.Load();
        IPresentationSourceService presentationSource = new OpenXmlPresentationSourceService(settings.ToReadLimits());
        IPresentationComparisonService comparisonService = new TextPresentationComparisonService();
        ISettingsDialogService settingsDialog = new WpfSettingsDialogService();
        IAboutDialogService aboutDialog = new WpfAboutDialogService(_diagnostics);

        var window = new MainWindow
        {
            DataContext = new MainWindowViewModel(
                filePicker,
                presentationSource,
                comparisonService,
                renderer,
                _diagnostics,
                settingsService,
                settingsDialog,
                aboutDialog,
                settings)
        };

        MainWindow = window;
        ShutdownMode = ShutdownMode.OnMainWindowClose;
        window.Show();
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        _diagnostics.RecordException("DispatcherUnhandledException", e.Exception);
        e.Handled = true;
        MessageBox.Show(
            "PptCompare encountered an unexpected error and must close. No presentation was modified.",
            "PptCompare",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        Shutdown(-1);
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            _diagnostics.RecordException("UnhandledException", exception);
        }
        else
        {
            _diagnostics.RecordEvent("UnhandledException", "type=Unknown");
        }
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        _diagnostics.RecordException("UnobservedTaskException", e.Exception);
        e.SetObserved();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        DispatcherUnhandledException -= OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
        TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;
        _diagnostics.RecordEvent("ApplicationStopped", $"exitCode={e.ApplicationExitCode}");
        base.OnExit(e);
    }

    [SuppressMessage(
        "Design",
        "CA1031:Do not catch general exception types",
        Justification = "Diagnostics are optional and must never prevent the application from starting.")]
    private static IApplicationDiagnostics CreateDiagnostics()
    {
        try
        {
            return new FileApplicationDiagnostics();
        }
        catch (Exception exception)
        {
            Debug.WriteLine($"PptCompare diagnostics were unavailable: {exception.GetType().Name}");
            return NullApplicationDiagnostics.Instance;
        }
    }
}
