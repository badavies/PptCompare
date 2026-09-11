using System.Windows;

namespace PptCompare.Services;

public interface IAboutDialogService
{
    void ShowAbout();
}

public sealed class WpfAboutDialogService(IApplicationDiagnostics diagnostics) : IAboutDialogService
{
    public void ShowAbout()
    {
        var window = new AboutWindow(diagnostics)
        {
            Owner = Application.Current?.MainWindow
        };
        window.ShowDialog();
    }
}
