using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PptCompare.Models;

namespace PptCompare.Controls;

public sealed class SlidePreviewControl : FrameworkElement
{
    private const long MaxRenderedImageBytes = 50L * 1024 * 1024;
    private static readonly Brush SlideBackground = CreateBrush(255, 255, 255);
    private static readonly Brush PlaceholderFill = CreateBrush(241, 245, 249);
    private static readonly Brush PlaceholderStroke = CreateBrush(148, 163, 184);
    private static readonly Brush TextFill = CreateBrush(248, 250, 252);
    private static readonly Brush TextForeground = CreateBrush(30, 41, 59);
    private static readonly Brush AddedBrush = CreateBrush(21, 128, 61);
    private static readonly Brush RemovedBrush = CreateBrush(194, 65, 59);
    private static readonly Brush ModifiedBrush = CreateBrush(202, 124, 22);
    private readonly Dictionary<string, ImageSource?> _embeddedImages = new(StringComparer.Ordinal);
    private string? _loadedRenderedPath;
    private ImageSource? _renderedImage;

    public static readonly DependencyProperty SlideProperty = DependencyProperty.Register(
        nameof(Slide),
        typeof(PresentationSlide),
        typeof(SlidePreviewControl),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender, OnSlideChanged));

    public static readonly DependencyProperty ChangesProperty = DependencyProperty.Register(
        nameof(Changes),
        typeof(IReadOnlyList<SlideElementChange>),
        typeof(SlidePreviewControl),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty IsRightSideProperty = DependencyProperty.Register(
        nameof(IsRightSide),
        typeof(bool),
        typeof(SlidePreviewControl),
        new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

    public PresentationSlide? Slide
    {
        get => (PresentationSlide?)GetValue(SlideProperty);
        set => SetValue(SlideProperty, value);
    }

    public IReadOnlyList<SlideElementChange>? Changes
    {
        get => (IReadOnlyList<SlideElementChange>?)GetValue(ChangesProperty);
        set => SetValue(ChangesProperty, value);
    }

    public bool IsRightSide
    {
        get => (bool)GetValue(IsRightSideProperty);
        set => SetValue(IsRightSideProperty, value);
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        base.OnRender(drawingContext);
        var slide = Slide;
        if (slide is null || ActualWidth <= 0 || ActualHeight <= 0)
        {
            DrawEmptyState(drawingContext);
            return;
        }

        var slideWidth = Math.Max(1, slide.Width);
        var slideHeight = Math.Max(1, slide.Height);
        var scale = Math.Min(ActualWidth / slideWidth, ActualHeight / slideHeight);
        var renderWidth = slideWidth * scale;
        var renderHeight = slideHeight * scale;
        var originX = (ActualWidth - renderWidth) / 2;
        var originY = (ActualHeight - renderHeight) / 2;
        var slideRect = new Rect(originX, originY, renderWidth, renderHeight);

        drawingContext.DrawRectangle(SlideBackground, new Pen(PlaceholderStroke, 1), slideRect);
        var renderedImage = TryLoadRenderedImage(slide.RenderedImagePath);
        if (renderedImage is not null)
        {
            drawingContext.DrawImage(renderedImage, slideRect);
        }
        else
        {
            DrawExtractedElements(drawingContext, slide, slideRect, scale);
        }

        DrawChangeOverlays(drawingContext, slideRect, scale);
    }

    private void DrawExtractedElements(
        DrawingContext drawingContext,
        PresentationSlide slide,
        Rect slideRect,
        double scale)
    {
        foreach (var element in slide.Elements)
        {
            var bounds = ToRenderRect(element.Bounds, slideRect, scale);
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                continue;
            }

            if (element.Kind == SlideElementKind.Image && element.ImageBytes is { Length: > 0 })
            {
                var image = GetEmbeddedImage(element);
                if (image is not null)
                {
                    drawingContext.DrawImage(image, bounds);
                    continue;
                }
            }

            var fill = element.Kind == SlideElementKind.Text ? TextFill : PlaceholderFill;
            drawingContext.DrawRoundedRectangle(
                fill,
                new Pen(PlaceholderStroke, 1),
                bounds,
                2,
                2);

            if (!string.IsNullOrWhiteSpace(element.Text) && bounds.Width >= 20 && bounds.Height >= 12)
            {
                var text = element.Text.Length > 300 ? element.Text[..300] + "…" : element.Text;
                var formatted = new FormattedText(
                    text,
                    CultureInfo.CurrentUICulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Segoe UI"),
                    Math.Clamp(14 * scale * 12_192_000 / Math.Max(1, slide.Width), 7, 16),
                    TextForeground,
                    VisualTreeHelper.GetDpi(this).PixelsPerDip)
                {
                    MaxTextWidth = Math.Max(1, bounds.Width - 8),
                    MaxTextHeight = Math.Max(1, bounds.Height - 8),
                    Trimming = TextTrimming.CharacterEllipsis
                };
                drawingContext.DrawText(formatted, new Point(bounds.X + 4, bounds.Y + 4));
            }
        }
    }

    private void DrawChangeOverlays(DrawingContext drawingContext, Rect slideRect, double scale)
    {
        if (Changes is not { Count: > 0 })
        {
            return;
        }

        foreach (var change in Changes)
        {
            var sourceBounds = IsRightSide ? change.RightBounds : change.LeftBounds;
            if (sourceBounds is null)
            {
                continue;
            }

            var bounds = ToRenderRect(sourceBounds, slideRect, scale);
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                continue;
            }

            var brush = change.Kind switch
            {
                SlideElementChangeKind.Added => AddedBrush,
                SlideElementChangeKind.Removed => RemovedBrush,
                SlideElementChangeKind.Replaced when IsRightSide => AddedBrush,
                SlideElementChangeKind.Replaced => RemovedBrush,
                _ => ModifiedBrush
            };
            var fill = brush.Clone();
            fill.Opacity = 0.10;
            fill.Freeze();
            drawingContext.DrawRectangle(fill, new Pen(brush, 3), bounds);
        }
    }

    private ImageSource? TryLoadRenderedImage(string? path)
    {
        if (string.Equals(path, _loadedRenderedPath, StringComparison.OrdinalIgnoreCase))
        {
            return _renderedImage;
        }

        _loadedRenderedPath = path;
        _renderedImage = null;
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            var file = new FileInfo(path);
            if (!file.Exists || file.Length <= 0 || file.Length > MaxRenderedImageBytes)
            {
                return null;
            }

            using var stream = new FileStream(
                file.FullName,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read | FileShare.Delete);
            _renderedImage = LoadBitmap(stream);
        }
        catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            _renderedImage = null;
        }

        return _renderedImage;
    }

    private ImageSource? GetEmbeddedImage(SlideElement element)
    {
        if (_embeddedImages.TryGetValue(element.ContentHash, out var cached))
        {
            return cached;
        }

        ImageSource? image = null;
        try
        {
            using var stream = new MemoryStream(element.ImageBytes!, false);
            image = LoadBitmap(stream);
        }
        catch (Exception exception) when (
            exception is IOException or NotSupportedException)
        {
        }

        _embeddedImages[element.ContentHash] = image;
        return image;
    }

    private static BitmapImage LoadBitmap(Stream stream)
    {
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.CreateOptions = BitmapCreateOptions.IgnoreColorProfile;
        image.DecodePixelWidth = 1600;
        image.StreamSource = stream;
        image.EndInit();
        image.Freeze();
        return image;
    }

    private static Rect ToRenderRect(SlideBounds bounds, Rect slideRect, double scale) =>
        new(
            slideRect.X + bounds.X * scale,
            slideRect.Y + bounds.Y * scale,
            Math.Max(0, bounds.Width * scale),
            Math.Max(0, bounds.Height * scale));

    private void DrawEmptyState(DrawingContext drawingContext)
    {
        drawingContext.DrawRectangle(PlaceholderFill, null, new Rect(RenderSize));
        if (ActualWidth < 80 || ActualHeight < 20)
        {
            return;
        }

        var text = new FormattedText(
            "No slide available",
            CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            new Typeface("Segoe UI"),
            13,
            PlaceholderStroke,
            VisualTreeHelper.GetDpi(this).PixelsPerDip);
        drawingContext.DrawText(
            text,
            new Point((ActualWidth - text.Width) / 2, (ActualHeight - text.Height) / 2));
    }

    private static void OnSlideChanged(DependencyObject sender, DependencyPropertyChangedEventArgs args)
    {
        var control = (SlidePreviewControl)sender;
        control._loadedRenderedPath = null;
        control._renderedImage = null;
        control._embeddedImages.Clear();
    }

    private static SolidColorBrush CreateBrush(byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }
}
