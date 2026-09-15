using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Threading;
using PptCompare.Controls;
using PptCompare.Models;

namespace PptCompare.Tests;

[TestClass]
public sealed class DiffTextBlockTests
{
    [TestMethod]
    public async Task SynchronousContentAndFontChangesProduceOneFinalDocument()
    {
        await RunOnStaThreadAsync(() =>
        {
            var control = new DiffTextBlock();
            var initialDocument = control.Document;
            var textChanges = 0;
            control.TextChanged += (_, _) => textChanges++;

            control.PlainText = "Stale text";
            control.Segments = [new DiffSegment("Stale text", DiffKind.Removed)];
            control.RichContent = CreateRichContent("Stale text");
            control.FontSize = 18;
            control.PlainText = "Final text";
            control.Segments =
            [
                new DiffSegment("Final ", DiffKind.Unchanged),
                new DiffSegment("text", DiffKind.Added)
            ];
            control.RichContent = CreateRichContent("Final text");
            control.FontSize = 24;

            Assert.AreSame(initialDocument, control.Document);
            Assert.AreEqual(0, textChanges);

            DrainDispatcherPastRender();

            Assert.AreNotSame(initialDocument, control.Document);
            Assert.AreEqual(1, textChanges);
            Assert.AreEqual(24, control.Document.FontSize);
            Assert.AreEqual("Final text", GetDocumentText(control.Document));
            var paragraph = Assert.IsInstanceOfType<Paragraph>(
                control.Document.Blocks.FirstBlock);
            Assert.IsTrue(paragraph.Inlines.OfType<Run>().All(
                run => run.FontWeight == FontWeights.Bold));
        });
    }

    [TestMethod]
    public async Task LaterChangeBurstSchedulesAnotherSingleRebuild()
    {
        await RunOnStaThreadAsync(() =>
        {
            var control = new DiffTextBlock
            {
                PlainText = "First"
            };
            DrainDispatcherPastRender();
            var firstDocument = control.Document;

            control.PlainText = "Intermediate";
            control.PlainText = "Second";
            control.Segments = [new DiffSegment("Second", DiffKind.Added)];

            Assert.AreSame(firstDocument, control.Document);

            DrainDispatcherPastRender();

            Assert.AreNotSame(firstDocument, control.Document);
            Assert.AreEqual("Second", GetDocumentText(control.Document));
        });
    }

    [TestMethod]
    public async Task HorizontalSpanRendersWithoutContinuationPlaceholderCell()
    {
        await RunOnStaThreadAsync(() =>
        {
            var source = new SlideTextTableBlock(
                [
                    new SlideTextTableRow(
                    [
                        new SlideTextTableCell(
                            [CreateParagraph("Merged"), CreateParagraph("Continuation")],
                            ColumnSpan: 2),
                        new SlideTextTableCell([CreateParagraph("Right")])
                    ]),
                    new SlideTextTableRow(
                    [
                        new SlideTextTableCell([CreateParagraph("One")]),
                        new SlideTextTableCell([CreateParagraph("Two")]),
                        new SlideTextTableCell([CreateParagraph("Three")])
                    ])
                ],
                3);

            var table = RenderTable(source);

            Assert.AreEqual(3, table.Columns.Count);
            var rows = table.RowGroups.Single().Rows;
            Assert.AreEqual(2, rows[0].Cells.Count);
            Assert.AreEqual(2, rows[0].Cells[0].ColumnSpan);
            Assert.AreEqual("Merged\r\nContinuation", GetCellText(rows[0].Cells[0]));
            Assert.AreEqual("Right", GetCellText(rows[0].Cells[1]));
        });
    }

    [TestMethod]
    public async Task VerticalSpanRendersWithoutContinuationPlaceholderCell()
    {
        await RunOnStaThreadAsync(() =>
        {
            var source = new SlideTextTableBlock(
                [
                    new SlideTextTableRow(
                    [
                        new SlideTextTableCell(
                            [CreateParagraph("Merged"), CreateParagraph("Continuation")],
                            RowSpan: 2),
                        new SlideTextTableCell([CreateParagraph("First row")])
                    ]),
                    new SlideTextTableRow(
                        [new SlideTextTableCell([CreateParagraph("Second row")])])
                ],
                2);

            var table = RenderTable(source);

            Assert.AreEqual(2, table.Columns.Count);
            var rows = table.RowGroups.Single().Rows;
            Assert.AreEqual(2, rows[0].Cells.Count);
            Assert.AreEqual(2, rows[0].Cells[0].RowSpan);
            Assert.AreEqual(1, rows[1].Cells.Count);
            Assert.AreEqual("Second row", GetCellText(rows[1].Cells[0]));
        });
    }

    private static IReadOnlyList<SlideTextBlock> CreateRichContent(string text) =>
    [
        new SlideTextParagraphBlock(
            new SlideTextParagraph(
                [new SlideTextRun(text, new SlideTextStyle(Bold: true))]))
    ];

    private static SlideTextParagraph CreateParagraph(string text) =>
        new([new SlideTextRun(text, new SlideTextStyle())]);

    private static Table RenderTable(SlideTextTableBlock source)
    {
        var control = new DiffTextBlock
        {
            PlainText = string.Join(
                Environment.NewLine + Environment.NewLine,
                source.Rows
                    .SelectMany(row => row.Cells)
                    .SelectMany(cell => cell.Paragraphs)
                    .Select(paragraph => paragraph.Text)),
            RichContent = [source]
        };
        DrainDispatcherPastRender();
        return Assert.IsInstanceOfType<Table>(control.Document.Blocks.FirstBlock);
    }

    private static string GetCellText(TableCell cell) =>
        new TextRange(cell.ContentStart, cell.ContentEnd)
            .Text
            .TrimEnd('\r', '\n');

    private static string GetDocumentText(FlowDocument document) =>
        new TextRange(document.ContentStart, document.ContentEnd)
            .Text
            .TrimEnd('\r', '\n');

    private static void DrainDispatcherPastRender()
    {
        var frame = new DispatcherFrame();
        _ = Dispatcher.CurrentDispatcher.BeginInvoke(
            DispatcherPriority.Background,
            () => frame.Continue = false);
        Dispatcher.PushFrame(frame);
    }

    private static async Task RunOnStaThreadAsync(Action action)
    {
        var completion = new TaskCompletionSource<object?>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                action();
                completion.SetResult(null);
            }
            catch (Exception exception)
            {
                completion.SetException(exception);
            }
        })
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        await completion.Task.WaitAsync(TimeSpan.FromSeconds(10));
    }
}
