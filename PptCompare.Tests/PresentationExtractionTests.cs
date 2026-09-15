using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class PresentationExtractionTests
{
    private const string PresentationContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    [TestMethod]
    public async Task HorizontalMergeContinuationIsFoldedIntoSpanningCell()
    {
        var path = CreateTestPath("horizontal-merge.pptx");
        try
        {
            string[] rows =
            [
                CreateTableRow(
                    CreateTableCell("Merged heading", columnSpan: 2),
                    CreateTableCell("Continuation note", isHorizontalContinuation: true),
                    CreateTableCell("Right")),
                CreateTableRow(
                    CreateTableCell("One"),
                    CreateTableCell("Two"),
                    CreateTableCell("Three"))
            ];
            CreatePresentationPackage(path, CreateTableSlideXml(3, rows));

            var presentation = await LoadAsync(path);

            var table = presentation.Slides.Single().TextContent
                .OfType<SlideTextTableBlock>()
                .Single();
            Assert.AreEqual(3, table.ColumnCount);
            Assert.HasCount(2, table.Rows[0].Cells);
            Assert.AreEqual(2, table.Rows[0].Cells[0].ColumnSpan);
            Assert.AreEqual(
                "Merged heading|Continuation note",
                GetCellText(table.Rows[0].Cells[0]));
            Assert.AreEqual("Right", GetCellText(table.Rows[0].Cells[1]));
            Assert.HasCount(3, table.Rows[1].Cells);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task VerticalMergeContinuationIsFoldedIntoSpanningCell()
    {
        var path = CreateTestPath("vertical-merge.pptx");
        try
        {
            string[] rows =
            [
                CreateTableRow(
                    CreateTableCell("Merged heading", rowSpan: 2),
                    CreateTableCell("First row")),
                CreateTableRow(
                    CreateTableCell("Continuation note", isVerticalContinuation: true),
                    CreateTableCell("Second row"))
            ];
            CreatePresentationPackage(path, CreateTableSlideXml(2, rows));

            var presentation = await LoadAsync(path);

            var table = presentation.Slides.Single().TextContent
                .OfType<SlideTextTableBlock>()
                .Single();
            Assert.AreEqual(2, table.ColumnCount);
            Assert.HasCount(2, table.Rows[0].Cells);
            Assert.AreEqual(2, table.Rows[0].Cells[0].RowSpan);
            Assert.AreEqual(
                "Merged heading|Continuation note",
                GetCellText(table.Rows[0].Cells[0]));
            Assert.HasCount(1, table.Rows[1].Cells);
            Assert.AreEqual("Second row", GetCellText(table.Rows[1].Cells[0]));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task ExternalHyperlinkRelationshipLoadsNormally()
    {
        var path = CreateTestPath("external-hyperlink.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                CreateHyperlinkSlideXml(),
                CreateExternalHyperlinkRelationship());

            var presentation = await LoadAsync(path);

            Assert.HasCount(1, presentation.Slides);
            Assert.AreEqual("Open example", presentation.Slides[0].Title);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private async Task<LoadedPresentation> LoadAsync(string path)
    {
        var source = new OpenXmlPresentationSourceService();
        return await source.LoadAsync(
            new PresentationReference(path, Path.GetFileName(path)),
            TestContext.CancellationToken);
    }

    private static string GetCellText(SlideTextTableCell cell) =>
        string.Join('|', cell.Paragraphs.Select(paragraph => paragraph.Text));

    private static void CreatePresentationPackage(
        string path,
        string slideXml,
        string? slideRelationships = null)
    {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var package = new ZipArchive(stream, ZipArchiveMode.Create);
        WriteTextEntry(package, "[Content_Types].xml", CreateContentTypes());
        WriteTextEntry(package, "_rels/.rels", CreatePackageRelationships());
        WriteTextEntry(package, "ppt/presentation.xml", CreatePresentationXml());
        WriteTextEntry(
            package,
            "ppt/_rels/presentation.xml.rels",
            CreatePresentationRelationships());
        WriteTextEntry(package, "ppt/slides/slide1.xml", slideXml);
        if (slideRelationships is not null)
        {
            WriteTextEntry(
                package,
                "ppt/slides/_rels/slide1.xml.rels",
                slideRelationships);
        }
    }

    private static string CreateTableSlideXml(int columnCount, IReadOnlyList<string> rows)
    {
        var gridColumns = string.Concat(
            Enumerable.Repeat("<a:gridCol w=\"2500000\"/>", columnCount));
        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                   <p:cSld><p:spTree>
                     <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                     <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                     {{CreateTitleShapeXml()}}
                     <p:graphicFrame>
                       <p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
                       <p:xfrm><a:off x="500000" y="1500000"/><a:ext cx="7500000" cy="3000000"/></p:xfrm>
                       <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                         <a:tbl>
                           <a:tblPr firstRow="1" bandRow="1"/>
                           <a:tblGrid>{{gridColumns}}</a:tblGrid>
                           {{string.Concat(rows)}}
                         </a:tbl>
                       </a:graphicData></a:graphic>
                     </p:graphicFrame>
                   </p:spTree></p:cSld>
                   <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
                 </p:sld>
                 """;
    }

    private static string CreateTitleShapeXml() =>
        """
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="500000" y="250000"/><a:ext cx="7500000" cy="750000"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Table title</a:t></a:r></a:p></p:txBody>
        </p:sp>
        """;

    private static string CreateTableRow(params string[] cells) =>
        $"<a:tr h=\"1000000\">{string.Concat(cells)}</a:tr>";

    private static string CreateTableCell(
        string text,
        int columnSpan = 1,
        int rowSpan = 1,
        bool isHorizontalContinuation = false,
        bool isVerticalContinuation = false)
    {
        var attributes = new StringBuilder();
        if (columnSpan > 1)
        {
            attributes
                .Append(" gridSpan=\"")
                .Append(columnSpan.ToString(CultureInfo.InvariantCulture))
                .Append('"');
        }

        if (rowSpan > 1)
        {
            attributes
                .Append(" rowSpan=\"")
                .Append(rowSpan.ToString(CultureInfo.InvariantCulture))
                .Append('"');
        }

        if (isHorizontalContinuation)
        {
            attributes.Append(" hMerge=\"1\"");
        }

        if (isVerticalContinuation)
        {
            attributes.Append(" vMerge=\"1\"");
        }

        return $$"""
                 <a:tc{{attributes}}>
                   <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>{{text}}</a:t></a:r></a:p></a:txBody>
                   <a:tcPr/>
                 </a:tc>
                 """;
    }

    private static string CreateHyperlinkSlideXml() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Linked text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr/>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rIdHyperlink"/></a:rPr><a:t>Open example</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
          <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
        </p:sld>
        """;

    private static string CreateExternalHyperlinkRelationship() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdHyperlink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/pptcompare" TargetMode="External"/>
        </Relationships>
        """;

    private static string CreateContentTypes() =>
        $$"""
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/ppt/presentation.xml" ContentType="{{PresentationContentType}}"/>
            <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
          </Types>
          """;

    private static string CreatePackageRelationships() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
        </Relationships>
        """;

    private static string CreatePresentationXml() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
          <p:sldSz cx="12192000" cy="6858000"/>
          <p:notesSz cx="6858000" cy="9144000"/>
        </p:presentation>
        """;

    private static string CreatePresentationRelationships() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
        </Relationships>
        """;

    private static void WriteTextEntry(ZipArchive package, string entryName, string contents)
    {
        var entry = package.CreateEntry(entryName);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(contents);
    }

    private static string CreateTestPath(string fileName)
    {
        var folder = Path.Combine(Path.GetTempPath(), "PptCompareTests");
        Directory.CreateDirectory(folder);
        return Path.Combine(folder, $"{Guid.NewGuid():N}-{fileName}");
    }

    public TestContext TestContext { get; set; } = null!;
}
