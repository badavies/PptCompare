using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PptCompare.Models;

namespace PptCompare.Controls;

public sealed class DiffTextBlock : RichTextBox
{
    private static readonly Brush AddedBackground = CreateBrush(0xD9, 0xF9, 0xE2);
    private static readonly Brush AddedForeground = CreateBrush(0x0F, 0x6B, 0x32);
    private static readonly Brush RemovedBackground = CreateBrush(0xFD, 0xE2, 0xE2);
    private static readonly Brush RemovedForeground = CreateBrush(0xA7, 0x2D, 0x2D);
    private static readonly Brush TableBorder = CreateBrush(0xCB, 0xD5, 0xE1);
    private static readonly Thickness NoThickness = new(0);

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

    public static readonly DependencyProperty RichContentProperty = DependencyProperty.Register(
        nameof(RichContent),
        typeof(IReadOnlyList<SlideTextBlock>),
        typeof(DiffTextBlock),
        new FrameworkPropertyMetadata(null, OnContentChanged));

    public DiffTextBlock()
    {
        IsReadOnly = true;
        IsUndoEnabled = false;
        IsDocumentEnabled = true;
        IsInactiveSelectionHighlightEnabled = true;
        IsReadOnlyCaretVisible = false;
        AcceptsReturn = true;
        BorderThickness = NoThickness;
        Padding = NoThickness;
        Background = Brushes.Transparent;
        HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
        VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
        VerticalContentAlignment = VerticalAlignment.Top;
    }

    public string PlainText
    {
        get => (string?)GetValue(PlainTextProperty) ?? string.Empty;
        set => SetValue(PlainTextProperty, value);
    }

    public IReadOnlyList<DiffSegment>? Segments
    {
        get => (IReadOnlyList<DiffSegment>?)GetValue(SegmentsProperty);
        set => SetValue(SegmentsProperty, value);
    }

    public IReadOnlyList<SlideTextBlock>? RichContent
    {
        get => (IReadOnlyList<SlideTextBlock>?)GetValue(RichContentProperty);
        set => SetValue(RichContentProperty, value);
    }

    protected override void OnPropertyChanged(DependencyPropertyChangedEventArgs e)
    {
        base.OnPropertyChanged(e);
        if (e.Property == FontSizeProperty)
        {
            RebuildDocument();
        }
    }

    private static void OnContentChanged(DependencyObject sender, DependencyPropertyChangedEventArgs args) =>
        ((DiffTextBlock)sender).RebuildDocument();

    private void RebuildDocument()
    {
        var document = CreateDocument();
        var content = RichContent;
        if (content is not { Count: > 0 } || !ContentMatchesPlainText(content))
        {
            AddPlainContent(document);
            Document = document;
            return;
        }

        var segmentText = Segments is { Count: > 0 }
            ? string.Concat(Segments.Select(segment => segment.Text))
            : PlainText;
        var cursor = new DiffCursor(
            string.Equals(segmentText, PlainText, StringComparison.Ordinal) ? Segments : null);
        var totalParagraphs = EnumerateParagraphs(content).Count();
        var paragraphIndex = 0;

        foreach (var block in content)
        {
            switch (block)
            {
                case SlideTextParagraphBlock paragraphBlock:
                    document.Blocks.Add(BuildParagraph(paragraphBlock.Paragraph, cursor));
                    AdvanceParagraph(cursor, ref paragraphIndex, totalParagraphs);
                    break;
                case SlideTextTableBlock tableBlock:
                    document.Blocks.Add(BuildTable(
                        tableBlock,
                        cursor,
                        ref paragraphIndex,
                        totalParagraphs));
                    break;
            }
        }

        Document = document;
    }

    private FlowDocument CreateDocument() => new()
    {
        PagePadding = NoThickness,
        ColumnWidth = double.PositiveInfinity,
        FontFamily = FontFamily,
        FontSize = FontSize,
        FontStretch = FontStretch,
        FontStyle = FontStyle,
        FontWeight = FontWeight,
        Foreground = Foreground
    };

    private bool ContentMatchesPlainText(IReadOnlyList<SlideTextBlock> content)
    {
        var richText = string.Join(
            Environment.NewLine + Environment.NewLine,
            EnumerateParagraphs(content).Select(paragraph => paragraph.Text));
        return string.Equals(richText, PlainText, StringComparison.Ordinal);
    }

    private void AddPlainContent(FlowDocument document)
    {
        var paragraph = CreateParagraph();
        if (Segments is not { Count: > 0 })
        {
            paragraph.Inlines.Add(new Run(PlainText));
        }
        else
        {
            foreach (var segment in Segments)
            {
                var run = new Run(segment.Text);
                ApplyDiffStyle(run, segment.Kind, null);
                paragraph.Inlines.Add(run);
            }
        }

        document.Blocks.Add(paragraph);
    }

    private static Table BuildTable(
        SlideTextTableBlock source,
        DiffCursor cursor,
        ref int paragraphIndex,
        int totalParagraphs)
    {
        var columnCount = Math.Clamp(source.ColumnCount, 1, 100);
        var table = new Table
        {
            CellSpacing = 0,
            Margin = new Thickness(0, 4, 0, 10)
        };
        for (var index = 0; index < columnCount; index++)
        {
            table.Columns.Add(new TableColumn());
        }

        var rowGroup = new TableRowGroup();
        table.RowGroups.Add(rowGroup);
        for (var rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++)
        {
            var sourceRow = source.Rows[rowIndex];
            var row = new TableRow();
            rowGroup.Rows.Add(row);
            var consumedColumns = 0;
            foreach (var sourceCell in sourceRow.Cells)
            {
                if (consumedColumns >= columnCount)
                {
                    break;
                }

                var columnSpan = Math.Clamp(sourceCell.ColumnSpan, 1, columnCount - consumedColumns);
                var rowSpan = Math.Clamp(sourceCell.RowSpan, 1, source.Rows.Count - rowIndex);
                var cell = new TableCell
                {
                    BorderBrush = TableBorder,
                    BorderThickness = new Thickness(0.5),
                    Padding = new Thickness(6, 4, 6, 4),
                    ColumnSpan = columnSpan,
                    RowSpan = rowSpan
                };
                row.Cells.Add(cell);
                consumedColumns += columnSpan;

                if (sourceCell.Paragraphs.Count == 0)
                {
                    cell.Blocks.Add(CreateParagraph());
                    continue;
                }

                foreach (var sourceParagraph in sourceCell.Paragraphs)
                {
                    cell.Blocks.Add(BuildParagraph(sourceParagraph, cursor));
                    AdvanceParagraph(cursor, ref paragraphIndex, totalParagraphs);
                }
            }
        }

        return table;
    }

    private static Paragraph BuildParagraph(SlideTextParagraph source, DiffCursor cursor)
    {
        var paragraph = CreateParagraph();
        foreach (var sourceRun in source.Runs)
        {
            foreach (var part in cursor.Consume(sourceRun.Text))
            {
                var run = new Run(part.Text);
                ApplySourceStyle(run, sourceRun.Style);
                ApplyDiffStyle(run, part.Kind, sourceRun.Style);
                paragraph.Inlines.Add(run);
            }
        }

        return paragraph;
    }

    private static Paragraph CreateParagraph() => new()
    {
        Margin = NoThickness,
        Padding = NoThickness
    };

    private static void AdvanceParagraph(
        DiffCursor cursor,
        ref int paragraphIndex,
        int totalParagraphs)
    {
        paragraphIndex++;
        if (paragraphIndex < totalParagraphs)
        {
            cursor.Skip(Environment.NewLine.Length * 2);
        }
    }

    private static void ApplySourceStyle(Run run, SlideTextStyle style)
    {
        if (style.Bold is not null)
        {
            run.FontWeight = style.Bold.Value ? FontWeights.Bold : FontWeights.Normal;
        }

        if (style.Italic is not null)
        {
            run.FontStyle = style.Italic.Value ? FontStyles.Italic : FontStyles.Normal;
        }

        var decorations = new TextDecorationCollection();
        if (style.Underline == true)
        {
            decorations.Add(TextDecorations.Underline[0]);
        }

        if (style.StrikeThrough == true)
        {
            decorations.Add(TextDecorations.Strikethrough[0]);
        }

        if (decorations.Count > 0)
        {
            run.TextDecorations = decorations;
        }

        if (TryCreateColorBrush(style.Color) is { } foreground)
        {
            run.Foreground = foreground;
        }

        if (!string.IsNullOrWhiteSpace(style.FontFamily) && style.FontFamily.Length <= 100)
        {
            try
            {
                run.FontFamily = new FontFamily(style.FontFamily);
            }
            catch (ArgumentException)
            {
                // Use the comparison pane font when the source font name is invalid.
            }
        }

        if (style.Baseline != SlideTextBaseline.Normal)
        {
            run.BaselineAlignment = style.Baseline == SlideTextBaseline.Superscript
                ? BaselineAlignment.Superscript
                : BaselineAlignment.Subscript;
        }
    }

    private static void ApplyDiffStyle(Run run, DiffKind kind, SlideTextStyle? sourceStyle)
    {
        switch (kind)
        {
            case DiffKind.Added:
                run.Background = AddedBackground;
                if (sourceStyle?.Color is null)
                {
                    run.Foreground = AddedForeground;
                }

                AddDecoration(run, TextDecorations.Underline[0]);
                break;
            case DiffKind.Removed:
                run.Background = RemovedBackground;
                if (sourceStyle?.Color is null)
                {
                    run.Foreground = RemovedForeground;
                }

                AddDecoration(run, TextDecorations.Strikethrough[0]);
                break;
            case DiffKind.Unchanged:
            default:
                break;
        }
    }

    private static void AddDecoration(Run run, TextDecoration decoration)
    {
        var decorations = run.TextDecorations is null
            ? new TextDecorationCollection()
            : new TextDecorationCollection(run.TextDecorations);
        if (!decorations.Contains(decoration))
        {
            decorations.Add(decoration);
        }

        run.TextDecorations = decorations;
    }

    private static IEnumerable<SlideTextParagraph> EnumerateParagraphs(
        IEnumerable<SlideTextBlock> blocks)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case SlideTextParagraphBlock paragraphBlock:
                    yield return paragraphBlock.Paragraph;
                    break;
                case SlideTextTableBlock table:
                    foreach (var paragraph in table.Rows
                                 .SelectMany(row => row.Cells)
                                 .SelectMany(cell => cell.Paragraphs))
                    {
                        yield return paragraph;
                    }

                    break;
            }
        }
    }

    private static SolidColorBrush? TryCreateColorBrush(string? value)
    {
        if (value is not ['#', _, _, _, _, _, _] ||
            !byte.TryParse(value.AsSpan(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var red) ||
            !byte.TryParse(value.AsSpan(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var green) ||
            !byte.TryParse(value.AsSpan(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var blue))
        {
            return null;
        }

        return CreateBrush(red, green, blue);
    }

    private static SolidColorBrush CreateBrush(byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }

    private sealed class DiffCursor(IReadOnlyList<DiffSegment>? segments)
    {
        private readonly IReadOnlyList<DiffSegment> _segments = segments ?? [];
        private int _segmentIndex;
        private int _offset;

        public IEnumerable<(string Text, DiffKind Kind)> Consume(string text)
        {
            var consumed = 0;
            while (consumed < text.Length)
            {
                MovePastEmptySegments();
                if (_segmentIndex >= _segments.Count)
                {
                    yield return (text[consumed..], DiffKind.Unchanged);
                    yield break;
                }

                var segment = _segments[_segmentIndex];
                var available = segment.Text.Length - _offset;
                var count = Math.Min(available, text.Length - consumed);
                yield return (text.Substring(consumed, count), segment.Kind);
                consumed += count;
                _offset += count;
            }
        }

        public void Skip(int count)
        {
            var remaining = count;
            while (remaining > 0)
            {
                MovePastEmptySegments();
                if (_segmentIndex >= _segments.Count)
                {
                    return;
                }

                var available = _segments[_segmentIndex].Text.Length - _offset;
                var consumed = Math.Min(available, remaining);
                _offset += consumed;
                remaining -= consumed;
            }
        }

        private void MovePastEmptySegments()
        {
            while (_segmentIndex < _segments.Count &&
                   _offset >= _segments[_segmentIndex].Text.Length)
            {
                _segmentIndex++;
                _offset = 0;
            }
        }
    }
}
