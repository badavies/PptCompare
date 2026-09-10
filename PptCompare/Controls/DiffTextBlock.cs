using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PptCompare.Models;

namespace PptCompare.Controls;

public sealed class DiffTextBlock : TextBlock
{
    private static readonly Brush AddedBackground = CreateBrush(0xD9, 0xF9, 0xE2);
    private static readonly Brush AddedForeground = CreateBrush(0x0F, 0x6B, 0x32);
    private static readonly Brush RemovedBackground = CreateBrush(0xFD, 0xE2, 0xE2);
    private static readonly Brush RemovedForeground = CreateBrush(0xA7, 0x2D, 0x2D);

    public static readonly DependencyProperty PlainTextProperty = DependencyProperty.Register(
        nameof(PlainText),
        typeof(string),
        typeof(DiffTextBlock),
        new FrameworkPropertyMetadata(string.Empty, OnContentChanged));

    public static readonly DependencyProperty SegmentsProperty = DependencyProperty.Register(
        nameof(Segments),
        typeof(IReadOnlyList<DiffSegment>),
        typeof(DiffTextBlock),
        new FrameworkPropertyMetadata(null, OnContentChanged));

    public string PlainText
    {
        get => (string)GetValue(PlainTextProperty);
        set => SetValue(PlainTextProperty, value);
    }

    public IReadOnlyList<DiffSegment>? Segments
    {
        get => (IReadOnlyList<DiffSegment>?)GetValue(SegmentsProperty);
        set => SetValue(SegmentsProperty, value);
    }

    private static void OnContentChanged(DependencyObject sender, DependencyPropertyChangedEventArgs args) =>
        ((DiffTextBlock)sender).RebuildInlines();

    private void RebuildInlines()
    {
        Inlines.Clear();

        if (Segments is not { Count: > 0 })
        {
            Inlines.Add(new Run(PlainText ?? string.Empty));
            return;
        }

        foreach (var segment in Segments)
        {
            var run = new Run(segment.Text);
            switch (segment.Kind)
            {
                case DiffKind.Added:
                    run.Background = AddedBackground;
                    run.Foreground = AddedForeground;
                    run.TextDecorations = System.Windows.TextDecorations.Underline;
                    break;
                case DiffKind.Removed:
                    run.Background = RemovedBackground;
                    run.Foreground = RemovedForeground;
                    run.TextDecorations = System.Windows.TextDecorations.Strikethrough;
                    break;
            }

            Inlines.Add(run);
        }
    }

    private static SolidColorBrush CreateBrush(byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }
}
