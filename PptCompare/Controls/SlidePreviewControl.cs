using System.Buffers.Binary;
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
    private const int MaxImageHeaderProbeBytes = 128 * 1024;
    private static readonly Brush SlideBackground = CreateBrush(255, 255, 255);
    private static readonly Brush PlaceholderFill = CreateBrush(241, 245, 249);
    private static readonly Brush PlaceholderStroke = CreateBrush(148, 163, 184);
    private static readonly Brush TextFill = CreateBrush(248, 250, 252);
    private static readonly Brush TextForeground = CreateBrush(30, 41, 59);
    private static readonly Brush AddedBrush = CreateBrush(21, 128, 61);
    private static readonly Brush RemovedBrush = CreateBrush(194, 65, 59);
    private static readonly Brush ModifiedBrush = CreateBrush(202, 124, 22);
    private readonly Dictionary<string, ImageSource?> _embeddedImages = new(StringComparer.Ordinal);
    private long _embeddedDecodedPixels;
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

            if (element is { Kind: SlideElementKind.Image, ImageBytes.Length: > 0 })
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

            if (string.IsNullOrWhiteSpace(element.Text) ||
                bounds is not { Width: >= 20, Height: >= 12 })
            {
                continue;
            }

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
            if (!HasSafeImageHeader(stream, out var header))
            {
                return null;
            }

            _renderedImage = LoadBitmap(stream, header);
        }
        catch (Exception exception) when (
            IsExpectedImageLoadFailure(exception))
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

        var imageBytes = element.ImageBytes;
        if (imageBytes is null ||
            !RasterImageHeaderValidator.TryValidate(imageBytes, out var header))
        {
            _embeddedImages[element.ContentHash] = null;
            return null;
        }

        var expectedDecodedPixels = RasterImageHeaderValidator.CalculateDecodedPixelCount(header);
        if (!RasterImageHeaderValidator.CanRetainDecodedPixels(
                _embeddedDecodedPixels,
                expectedDecodedPixels))
        {
            _embeddedImages[element.ContentHash] = null;
            return null;
        }

        ImageSource? image = null;
        try
        {
            using var stream = new MemoryStream(imageBytes, false);
            var bitmap = LoadBitmap(stream, header);
            var actualDecodedPixels = (long)bitmap.PixelWidth * bitmap.PixelHeight;
            if (RasterImageHeaderValidator.TryReserveDecodedPixels(
                    ref _embeddedDecodedPixels,
                    actualDecodedPixels))
            {
                image = bitmap;
            }
        }
        catch (Exception exception) when (
            IsExpectedImageLoadFailure(exception))
        {
        }

        _embeddedImages[element.ContentHash] = image;
        return image;
    }

    private static bool HasSafeImageHeader(
        Stream stream,
        out RasterImageHeader header)
    {
        header = default;
        var initialPosition = stream.Position;
        try
        {
            Span<byte> shortHeader = stackalloc byte[24];
            var shortHeaderLength = stream.ReadAtLeast(
                shortHeader,
                shortHeader.Length,
                throwOnEndOfStream: false);
            if (RasterImageHeaderValidator.TryValidate(
                shortHeader[..shortHeaderLength],
                out header))
            {
                return true;
            }

            if (shortHeaderLength < 2 ||
                shortHeader[0] != 0xFF ||
                shortHeader[1] != 0xD8)
            {
                return false;
            }

            stream.Position = initialPosition;
            var bytesToRead = checked((int)Math.Min(
                MaxImageHeaderProbeBytes,
                Math.Max(0, stream.Length - initialPosition)));
            if (bytesToRead == 0)
            {
                return false;
            }

            var headerBytes = new byte[bytesToRead];
            var bytesRead = stream.ReadAtLeast(
                headerBytes,
                bytesToRead,
                throwOnEndOfStream: false);
            return RasterImageHeaderValidator.TryValidate(
                headerBytes.AsSpan(0, bytesRead),
                out header);
        }
        finally
        {
            stream.Position = initialPosition;
        }
    }

    private static bool IsExpectedImageLoadFailure(Exception exception) =>
        exception is ArgumentException or FormatException or IOException or
            InvalidOperationException or NotSupportedException or UnauthorizedAccessException;

    private static BitmapImage LoadBitmap(Stream stream, RasterImageHeader header)
    {
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.CreateOptions = BitmapCreateOptions.IgnoreColorProfile;
        if (header.PixelWidth >= header.PixelHeight)
        {
            image.DecodePixelWidth = Math.Min(
                RasterImageHeaderValidator.MaxDecodeDimension,
                header.PixelWidth);
        }
        else
        {
            image.DecodePixelHeight = Math.Min(
                RasterImageHeaderValidator.MaxDecodeDimension,
                header.PixelHeight);
        }

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
        control._embeddedDecodedPixels = 0;
    }

    private static SolidColorBrush CreateBrush(byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }
}

internal enum RasterImageFormat
{
    Png,
    Jpeg
}

internal readonly record struct RasterImageHeader(
    RasterImageFormat Format,
    int PixelWidth,
    int PixelHeight);

internal static class RasterImageHeaderValidator
{
    internal const int MaxSourceDimension = 32_768;
    private const long MaxSourcePixelCount = 40_000_000;
    internal const int MaxAspectRatio = 1_000;
    internal const int MaxDecodeDimension = 1_600;
    private const long MaxRetainedDecodedPixels = 20_000_000;

    private const int PngHeaderLength = 24;

    internal static bool TryValidate(
        ReadOnlySpan<byte> bytes,
        out RasterImageHeader header)
    {
        header = default;

        if (TryReadPngDimensions(bytes, out var width, out var height))
        {
            return TryCreateHeader(RasterImageFormat.Png, width, height, out header);
        }

        if (TryReadJpegDimensions(bytes, out width, out height))
        {
            return TryCreateHeader(RasterImageFormat.Jpeg, width, height, out header);
        }

        return false;
    }

    internal static long CalculateDecodedPixelCount(RasterImageHeader header)
    {
        var longestSide = Math.Max(header.PixelWidth, header.PixelHeight);
        var shortestSide = Math.Min(header.PixelWidth, header.PixelHeight);
        if (longestSide <= MaxDecodeDimension)
        {
            return (long)longestSide * shortestSide;
        }

        var decodedShortestSide = Math.Max(
            1,
            (((long)shortestSide * MaxDecodeDimension) + longestSide - 1) / longestSide);
        return MaxDecodeDimension * decodedShortestSide;
    }

    internal static bool CanRetainDecodedPixels(long retainedPixels, long candidatePixels) =>
        retainedPixels is >= 0 and <= MaxRetainedDecodedPixels &&
        candidatePixels > 0 &&
        candidatePixels <= MaxRetainedDecodedPixels - retainedPixels;

    internal static bool TryReserveDecodedPixels(
        ref long retainedPixels,
        long candidatePixels)
    {
        if (!CanRetainDecodedPixels(retainedPixels, candidatePixels))
        {
            return false;
        }

        retainedPixels += candidatePixels;
        return true;
    }

    private static bool TryReadPngDimensions(
        ReadOnlySpan<byte> bytes,
        out uint width,
        out uint height)
    {
        width = 0;
        height = 0;
        if (bytes.Length < PngHeaderLength ||
            bytes[0] != 0x89 ||
            bytes[1] != 0x50 ||
            bytes[2] != 0x4E ||
            bytes[3] != 0x47 ||
            bytes[4] != 0x0D ||
            bytes[5] != 0x0A ||
            bytes[6] != 0x1A ||
            bytes[7] != 0x0A ||
            BinaryPrimitives.ReadUInt32BigEndian(bytes[8..12]) != 13 ||
            bytes[12] != 0x49 ||
            bytes[13] != 0x48 ||
            bytes[14] != 0x44 ||
            bytes[15] != 0x52)
        {
            return false;
        }

        width = BinaryPrimitives.ReadUInt32BigEndian(bytes[16..20]);
        height = BinaryPrimitives.ReadUInt32BigEndian(bytes[20..24]);
        return true;
    }

    private static bool TryReadJpegDimensions(
        ReadOnlySpan<byte> bytes,
        out uint width,
        out uint height)
    {
        width = 0;
        height = 0;
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            return false;
        }

        var position = 2;
        while (position < bytes.Length)
        {
            if (bytes[position++] != 0xFF)
            {
                return false;
            }

            while (position < bytes.Length && bytes[position] == 0xFF)
            {
                position++;
            }

            if (position >= bytes.Length)
            {
                return false;
            }

            var marker = bytes[position++];
            if (marker == 0x00 || marker is 0xD8 or 0xD9 or 0xDA)
            {
                return false;
            }

            if (marker == 0x01 || marker is >= 0xD0 and <= 0xD7)
            {
                continue;
            }

            if (bytes.Length - position < 2)
            {
                return false;
            }

            var segmentLength = BinaryPrimitives.ReadUInt16BigEndian(bytes[position..]);
            if (segmentLength < 2 || segmentLength > bytes.Length - position)
            {
                return false;
            }

            if (IsStartOfFrameMarker(marker))
            {
                if (segmentLength < 8)
                {
                    return false;
                }

                var componentCount = bytes[position + 7];
                if (componentCount == 0 || segmentLength < 8 + (3 * componentCount))
                {
                    return false;
                }

                height = BinaryPrimitives.ReadUInt16BigEndian(bytes[(position + 3)..]);
                width = BinaryPrimitives.ReadUInt16BigEndian(bytes[(position + 5)..]);
                return true;
            }

            position += segmentLength;
        }

        return false;
    }

    private static bool IsStartOfFrameMarker(byte marker) =>
        marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or
            0xC5 or 0xC6 or 0xC7 or
            0xC9 or 0xCA or 0xCB or
            0xCD or 0xCE or 0xCF;

    private static bool TryCreateHeader(
        RasterImageFormat format,
        uint width,
        uint height,
        out RasterImageHeader header)
    {
        header = default;
        if (width == 0 ||
            height == 0 ||
            width > MaxSourceDimension ||
            height > MaxSourceDimension)
        {
            return false;
        }

        var pixelCount = (long)width * height;
        var smallerDimension = Math.Min(width, height);
        var largerDimension = Math.Max(width, height);
        if (pixelCount > MaxSourcePixelCount ||
            largerDimension > (long)smallerDimension * MaxAspectRatio)
        {
            return false;
        }

        header = new RasterImageHeader(format, checked((int)width), checked((int)height));
        return true;
    }
}
