using System.Windows;
using System.Windows.Controls;

namespace PptCompare;

public partial class MainWindow : Window
{
    private bool _synchronizingScroll;

    public MainWindow()
    {
        InitializeComponent();
    }

    protected override void OnClosed(EventArgs e)
    {
        if (DataContext is IDisposable disposable)
        {
            disposable.Dispose();
        }

        base.OnClosed(e);
    }

    private void ComparisonPane_OnScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (_synchronizingScroll || sender is not ScrollViewer source)
        {
            return;
        }

        var target = ReferenceEquals(source, LeftPane) ? RightPane : LeftPane;

        try
        {
            _synchronizingScroll = true;
            target.ScrollToHorizontalOffset(source.HorizontalOffset);
            target.ScrollToVerticalOffset(source.VerticalOffset);
        }
        finally
        {
            _synchronizingScroll = false;
        }
    }
}
