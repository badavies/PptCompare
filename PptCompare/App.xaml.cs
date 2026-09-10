using System.Windows;
using System.Windows.Threading;
using PptCompare.Services;
using PptCompare.ViewModels;

namespace PptCompare;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        IFilePickerService filePicker = new WpfFilePickerService();
        IPresentationSourceService presentationSource = new OpenXmlPresentationSourceService();
        IPresentationComparisonService comparisonService = new TextPresentationComparisonService();

        var window = new MainWindow
        {
            DataContext = new MainWindowViewModel(filePicker, presentationSource, comparisonService)
        };

        MainWindow = window;
        ShutdownMode = ShutdownMode.OnMainWindowClose;
        window.Show();
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        e.Handled = true;
        MessageBox.Show(
            "PptCompare encountered an unexpected error and must close. No presentation was modified.",
            "PptCompare",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        Shutdown(-1);
    }
}
