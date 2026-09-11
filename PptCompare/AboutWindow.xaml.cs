using System.Runtime.InteropServices;
using System.Windows;
using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare;

public partial class AboutWindow : Window
{
    private readonly IApplicationDiagnostics _diagnostics;

    public AboutWindow(IApplicationDiagnostics diagnostics)
    {
        ArgumentNullException.ThrowIfNull(diagnostics);
        InitializeComponent();
        _diagnostics = diagnostics;
        VersionText.Text = $"Version {ApplicationInfo.Version}";
        DetailVersionText.Text = ApplicationInfo.Version;
        PlatformText.Text = $"{RuntimeInformation.FrameworkDescription} · {RuntimeInformation.ProcessArchitecture}";
        DiagnosticsLocationText.Text = diagnostics.DisplayLogLocation;
    }

    private void CopyTechnicalDetails_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(_diagnostics.CreateSupportSummary());
            MessageBox.Show(
                this,
                "Technical details copied. They do not contain presentation names, paths or content.",
                "PptCompare",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (ExternalException exception)
        {
            _diagnostics.RecordException("ClipboardCopyFailed", exception);
            MessageBox.Show(
                this,
                "The clipboard is currently unavailable. Close any clipboard utility and try again.",
                "PptCompare",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }
}
