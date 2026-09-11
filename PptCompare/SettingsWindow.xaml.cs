using System.Windows;
using PptCompare.Models;
using PptCompare.ViewModels;

namespace PptCompare;

public partial class SettingsWindow : Window
{
    private readonly SettingsWindowViewModel _viewModel;

    public SettingsWindow(ApplicationSettings settings)
    {
        InitializeComponent();
        _viewModel = new SettingsWindowViewModel(settings);
        DataContext = _viewModel;
    }

    public ApplicationSettings? Result { get; private set; }

    private void Save_OnClick(object sender, RoutedEventArgs e)
    {
        if (!_viewModel.TryCreateSettings(out var settings, out var error))
        {
            MessageBox.Show(this, error, "Check settings", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Result = settings;
        DialogResult = true;
    }

    private void RestoreDefaults_OnClick(object sender, RoutedEventArgs e) =>
        _viewModel.RestoreDefaults();
}
