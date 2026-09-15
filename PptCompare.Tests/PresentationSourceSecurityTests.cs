using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using PptCompare.Models;
using PptCompare.Services;

namespace PptCompare.Tests;

[TestClass]
public sealed class PresentationSourceSecurityTests
{
    private const string PresentationContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string MacroPresentationContentType =
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
    private const string TemplateContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";

    [TestMethod]
    public async Task MissingBlankEmptyAndOversizedFilesAreRejected()
    {
        var missingPath = CreateTestPath("missing.pptx");
        await AssertRejectedAsync(string.Empty);
        await AssertRejectedAsync("   ");
        await AssertRejectedAsync(missingPath);

        var emptyPath = CreateTestPath("empty.pptx");
        var oversizedPath = CreateTestPath("oversized.pptx");
        try
        {
            await File.WriteAllBytesAsync(emptyPath, [], TestContext.CancellationToken);
            await File.WriteAllBytesAsync(oversizedPath, [1, 2], TestContext.CancellationToken);

            await AssertRejectedAsync(emptyPath);
            await AssertRejectedAsync(
                oversizedPath,
                new PresentationReadLimits(MaxFileBytes: 1));
        }
        finally
        {
            File.Delete(emptyPath);
            File.Delete(oversizedPath);
        }
    }

    [TestMethod]
    public async Task NonPowerPointExtensionIsRejected()
    {
        var path = CreateTestPath("not-a-presentation.txt");
        try
        {
            await File.WriteAllTextAsync(path, "plain text", TestContext.CancellationToken);

            await AssertRejectedAsync(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task CorruptOpenXmlPackageIsRejected()
    {
        var path = CreateTestPath("corrupt.pptx");
        try
        {
            await File.WriteAllTextAsync(
                path,
                "not an Open XML package",
                TestContext.CancellationToken);

            await AssertRejectedAsync(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task DocumentTypeMustMatchPowerPointExtension()
    {
        var templatePath = CreateTestPath("template-renamed.pptx");
        var renamedPresentationPath = CreateTestPath("presentation-renamed.pptm");
        try
        {
            CreatePresentationPackage(templatePath, TemplateContentType);
            CreatePresentationPackage(renamedPresentationPath, PresentationContentType);

            var templateException = await AssertRejectedAsync(templatePath);
            var macroException = await AssertRejectedAsync(renamedPresentationPath);

            StringAssert.Contains(templateException.Message, "document type");
            StringAssert.Contains(macroException.Message, "document type");
        }
        finally
        {
            File.Delete(templatePath);
            File.Delete(renamedPresentationPath);
        }
    }

    [TestMethod]
    public async Task ValidPresentationAndMacroPresentationAreAccepted()
    {
        var presentationPath = CreateTestPath("valid.pptx");
        var macroPresentationPath = CreateTestPath("valid.pptm");
        try
        {
            CreatePresentationPackage(presentationPath, PresentationContentType);
            CreatePresentationPackage(macroPresentationPath, MacroPresentationContentType);
            var source = new OpenXmlPresentationSourceService();

            var presentation = await source.LoadAsync(
                CreateReference(presentationPath),
                TestContext.CancellationToken);
            var macroPresentation = await source.LoadAsync(
                CreateReference(macroPresentationPath),
                TestContext.CancellationToken);

            Assert.HasCount(1, presentation.Slides);
            Assert.HasCount(1, macroPresentation.Slides);
        }
        finally
        {
            File.Delete(presentationPath);
            File.Delete(macroPresentationPath);
        }
    }

    [TestMethod]
    public async Task SlideXmlCharacterBudgetIsEnforcedDuringPartLoad()
    {
        var path = CreateTestPath("oversized-slide-xml.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                PresentationContentType,
                slideXml: CreateTextSlideXml(new string('A', 4_000)));

            var exception = await AssertRejectedAsync(
                path,
                new PresentationReadLimits(MaxCharactersPerPart: 2_000));

            Assert.IsInstanceOfType<XmlException>(exception.InnerException);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task DocumentTypeDeclarationInSlideXmlIsRejected()
    {
        var path = CreateTestPath("slide-dtd.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                PresentationContentType,
                slideXml: CreateTextSlideXml(
                    "&injected;",
                    "<!DOCTYPE p:sld [<!ENTITY injected \"untrusted text\">]>"));

            var exception = await AssertRejectedAsync(path);

            Assert.IsInstanceOfType<XmlException>(exception.InnerException);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task UnexpectedSlideRootIsRejectedAsInvalidPresentation()
    {
        var path = CreateTestPath("unexpected-slide-root.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                PresentationContentType,
                slideXml: CreateThemeXml());

            var exception = await AssertRejectedAsync(path);

            Assert.IsInstanceOfType<InvalidDataException>(exception.InnerException);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task StrictSlideNamespacesAreNormalizedBeforeExtraction()
    {
        var path = CreateTestPath("strict-slide.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                PresentationContentType,
                slideXml: CreateTextSlideXml("Strict text", useStrictNamespaces: true),
                useStrictPackage: true);
            var source = new OpenXmlPresentationSourceService();

            var loaded = await source.LoadAsync(
                CreateReference(path),
                TestContext.CancellationToken);

            var slide = loaded.Slides.Single();
            Assert.AreEqual("Strict text", slide.TitleContent?.Text);
            var textBlock = Assert.IsInstanceOfType<SlideTextParagraphBlock>(
                slide.TextContent.Single());
            Assert.AreEqual("Strict text", textBlock.Paragraph.Text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task SharedLayoutMasterAndThemePartsPreserveExtractedText()
    {
        var path = CreateTestPath("shared-presentation-parts.pptx");
        try
        {
            CreatePresentationWithSharedParts(path);
            var conversionCounts = new Dictionary<(PresentationXmlPartKind Kind, string Uri), int>();
            var source = new OpenXmlPresentationSourceService(
                null,
                (kind, uri) =>
                {
                    var key = (kind, uri.OriginalString);
                    conversionCounts.TryGetValue(key, out var count);
                    conversionCounts[key] = count + 1;
                });

            var loaded = await source.LoadAsync(
                CreateReference(path),
                TestContext.CancellationToken);

            Assert.HasCount(2, loaded.Slides);
            for (var index = 0; index < loaded.Slides.Count; index++)
            {
                var blocks = loaded.Slides[index].TextContent
                    .OfType<SlideTextParagraphBlock>()
                    .ToList();
                Assert.HasCount(2, blocks);
                Assert.AreEqual($"Title {index + 1}", blocks[0].Paragraph.Text);
                Assert.AreEqual($"Body {index + 1}", blocks[1].Paragraph.Text);
                Assert.AreEqual(
                    "#1A2B3C",
                    blocks[0].Paragraph.Runs.Single().Style.Color);
                Assert.AreEqual(
                    "#1A2B3C",
                    blocks[1].Paragraph.Runs.Single().Style.Color);
            }

            Assert.AreEqual(3, conversionCounts.Count);
            Assert.AreEqual(
                1,
                conversionCounts[(PresentationXmlPartKind.SlideLayout, "/ppt/slideLayouts/slideLayout1.xml")]);
            Assert.AreEqual(
                1,
                conversionCounts[(PresentationXmlPartKind.SlideMaster, "/ppt/slideMasters/slideMaster1.xml")]);
            Assert.AreEqual(
                1,
                conversionCounts[(PresentationXmlPartKind.Theme, "/ppt/theme/theme1.xml")]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task AggregateXmlBudgetCountsEachSharedPartOnceAndRejectsTheNextByte()
    {
        var path = CreateTestPath("aggregate-xml-budget.pptx");
        try
        {
            CreatePresentationWithSharedParts(path);
            var uniqueXmlBytes = GetSharedPresentationXmlByteCount();
            var source = new OpenXmlPresentationSourceService(
                new PresentationReadLimits(MaxTotalXmlPartBytes: uniqueXmlBytes));

            var loaded = await source.LoadAsync(
                CreateReference(path),
                TestContext.CancellationToken);
            var exception = await AssertRejectedAsync(
                path,
                new PresentationReadLimits(MaxTotalXmlPartBytes: uniqueXmlBytes - 1));

            Assert.HasCount(2, loaded.Slides);
            StringAssert.Contains(exception.Message, "too much XML data");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task MissingRequiredPresentationAndSlidePartsAreRejected()
    {
        var missingPresentationPath = CreateTestPath("missing-presentation.pptx");
        var missingSlidePath = CreateTestPath("missing-slide.pptx");
        try
        {
            CreatePresentationPackage(
                missingPresentationPath,
                PresentationContentType,
                includePresentationPart: false);
            CreatePresentationPackage(
                missingSlidePath,
                PresentationContentType,
                includeSlidePart: false);

            await AssertRejectedAsync(missingPresentationPath);
            await AssertRejectedAsync(missingSlidePath);
        }
        finally
        {
            File.Delete(missingPresentationPath);
            File.Delete(missingSlidePath);
        }
    }

    [TestMethod]
    public async Task RelationshipReferenceAndUniquePartBudgetsAreEnforced()
    {
        var repeatedReferencePath = CreateTestPath("too-many-references.pptx");
        var uniquePartsPath = CreateTestPath("too-many-parts.pptx");
        try
        {
            CreatePresentationPackage(
                repeatedReferencePath,
                PresentationContentType,
                [[1, 2, 3, 4]],
                [0, 0, 0]);
            CreatePresentationPackage(
                uniquePartsPath,
                PresentationContentType,
                [[1], [2]],
                [0, 1]);

            await AssertRejectedAsync(
                repeatedReferencePath,
                new PresentationReadLimits(
                    MaxRelationshipReferences: 2,
                    MaxUniqueRelatedParts: 1));
            await AssertRejectedAsync(
                uniquePartsPath,
                new PresentationReadLimits(
                    MaxRelationshipReferences: 2,
                    MaxUniqueRelatedParts: 1));
        }
        finally
        {
            File.Delete(repeatedReferencePath);
            File.Delete(uniquePartsPath);
        }
    }

    [TestMethod]
    public async Task CumulativeDecompressedPartBudgetIsEnforced()
    {
        var path = CreateTestPath("cumulative-parts.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                PresentationContentType,
                [[1, 2, 3, 4], [5, 6, 7, 8]],
                [0, 1]);

            await AssertRejectedAsync(
                path,
                new PresentationReadLimits(
                    MaxRelatedPartBytes: 4,
                    MaxRelationshipReferences: 2,
                    MaxUniqueRelatedParts: 2,
                    MaxTotalRelatedPartBytes: 6));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task RepeatedPartReferencesReuseHashAndPreviewAccounting()
    {
        var path = CreateTestPath("repeated-part.pptx");
        try
        {
            CreatePresentationPackage(
                path,
                PresentationContentType,
                [[1, 2, 3, 4]],
                [0, 0]);
            var source = new OpenXmlPresentationSourceService(
                new PresentationReadLimits(
                    MaxRelatedPartBytes: 4,
                    MaxRelationshipReferences: 2,
                    MaxUniqueRelatedParts: 1,
                    MaxTotalRelatedPartBytes: 4,
                    MaxPreviewImageBytes: 4,
                    MaxTotalPreviewImageBytes: 4));

            var loaded = await source.LoadAsync(
                CreateReference(path),
                TestContext.CancellationToken);

            var elements = loaded.Slides.Single().Elements;
            Assert.HasCount(2, elements);
            Assert.AreEqual(elements[0].ContentHash, elements[1].ContentHash);
            Assert.IsNotNull(elements[0].ImageBytes);
            Assert.AreSame(elements[0].ImageBytes, elements[1].ImageBytes);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public async Task LoadingAndComparisonObserveCancellation()
    {
        var path = CreateTestPath("cancel.pptx");
        try
        {
            CreatePresentationPackage(path, PresentationContentType);
            using var cancellation = new CancellationTokenSource();
            await cancellation.CancelAsync();
            var source = new OpenXmlPresentationSourceService();
            var comparison = new TextPresentationComparisonService();
            var presentation = CreateModelPresentation();

            await AssertCanceledAsync(() =>
                source.LoadAsync(CreateReference(path), cancellation.Token));
            await AssertCanceledAsync(() =>
                comparison.CompareAsync(presentation, presentation, cancellation.Token));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void InvalidOrInconsistentReadLimitsAreRejected()
    {
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
            new OpenXmlPresentationSourceService(
                new PresentationReadLimits(MaxRelationshipReferences: 0)));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
            new OpenXmlPresentationSourceService(
                new PresentationReadLimits(
                    MaxRelationshipReferences: 1,
                    MaxUniqueRelatedParts: 2)));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
            new OpenXmlPresentationSourceService(
                new PresentationReadLimits(
                    MaxRelatedPartBytes: 10,
                    MaxTotalRelatedPartBytes: 9)));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
            new OpenXmlPresentationSourceService(
                new PresentationReadLimits(
                    MaxPreviewImageBytes: 10,
                    MaxTotalPreviewImageBytes: 9)));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() =>
            new OpenXmlPresentationSourceService(
                new PresentationReadLimits(MaxTotalXmlPartBytes: 0)));
    }

    private async Task<PresentationLoadException> AssertRejectedAsync(
        string path,
        PresentationReadLimits? limits = null)
    {
        var source = new OpenXmlPresentationSourceService(limits);
        return await Assert.ThrowsExactlyAsync<PresentationLoadException>(() =>
            source.LoadAsync(CreateReference(path), TestContext.CancellationToken));
    }

    private static async Task AssertCanceledAsync(Func<Task> action)
    {
        try
        {
            await action();
        }
        catch (OperationCanceledException)
        {
            return;
        }

        Assert.Fail("The operation should have observed cancellation.");
    }

    private static PresentationReference CreateReference(string path) =>
        new(path, Path.GetFileName(path));

    private static LoadedPresentation CreateModelPresentation()
    {
        var slide = new PresentationSlide(
            1,
            "Title",
            ["Title", "Body"],
            12_192_000,
            6_858_000,
            []);
        return new LoadedPresentation(CreateReference("model.pptx"), [slide]);
    }

    private static void CreatePresentationPackage(
        string path,
        string mainContentType,
        IReadOnlyList<byte[]>? imageParts = null,
        IReadOnlyList<int>? picturePartIndexes = null,
        bool includePresentationPart = true,
        bool includeSlidePart = true,
        string? slideXml = null,
        bool useStrictPackage = false)
    {
        imageParts ??= [];
        picturePartIndexes ??= [];
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var package = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteTextEntry(
            package,
            "[Content_Types].xml",
            CreateContentTypes(mainContentType));
        WriteTextEntry(
            package,
            "_rels/.rels",
            CreatePackageRelationships(useStrictPackage));
        if (!includePresentationPart)
        {
            return;
        }

        WriteTextEntry(
            package,
            "ppt/presentation.xml",
            CreatePresentationXml(useStrictPackage));
        WriteTextEntry(
            package,
            "ppt/_rels/presentation.xml.rels",
            CreatePresentationRelationships(useStrictPackage));
        if (!includeSlidePart)
        {
            return;
        }

        WriteTextEntry(
            package,
            "ppt/slides/slide1.xml",
            slideXml ?? CreateSlideXml(picturePartIndexes));
        if (imageParts.Count > 0)
        {
            WriteTextEntry(
                package,
                "ppt/slides/_rels/slide1.xml.rels",
                CreateSlideRelationships(imageParts.Count));
        }

        for (var index = 0; index < imageParts.Count; index++)
        {
            var entry = package.CreateEntry($"ppt/media/image{index + 1}.png");
            using var entryStream = entry.Open();
            entryStream.Write(imageParts[index]);
        }
    }

    private static void CreatePresentationWithSharedParts(string path)
    {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var package = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteTextEntry(
            package,
            "[Content_Types].xml",
            CreateSharedPartContentTypes());
        WriteTextEntry(package, "_rels/.rels", CreatePackageRelationships());
        WriteTextEntry(package, "ppt/presentation.xml", CreateTwoSlidePresentationXml());
        WriteTextEntry(
            package,
            "ppt/_rels/presentation.xml.rels",
            CreateTwoSlidePresentationRelationships());
        for (var slideNumber = 1; slideNumber <= 2; slideNumber++)
        {
            WriteTextEntry(
                package,
                $"ppt/slides/slide{slideNumber}.xml",
                CreateInheritedTextSlideXml(slideNumber));
            WriteTextEntry(
                package,
                $"ppt/slides/_rels/slide{slideNumber}.xml.rels",
                CreateSlideLayoutRelationship());
        }

        WriteTextEntry(
            package,
            "ppt/slideLayouts/slideLayout1.xml",
            CreateSlideLayoutXml());
        WriteTextEntry(
            package,
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            CreateSlideMasterRelationship());
        WriteTextEntry(
            package,
            "ppt/slideMasters/slideMaster1.xml",
            CreateSlideMasterXml());
        WriteTextEntry(
            package,
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            CreateThemeRelationship());
        WriteTextEntry(package, "ppt/theme/theme1.xml", CreateThemeXml());
    }

    private static long GetSharedPresentationXmlByteCount() =>
        Encoding.UTF8.GetByteCount(CreateTwoSlidePresentationXml()) +
        Encoding.UTF8.GetByteCount(CreateInheritedTextSlideXml(1)) +
        Encoding.UTF8.GetByteCount(CreateInheritedTextSlideXml(2)) +
        Encoding.UTF8.GetByteCount(CreateSlideLayoutXml()) +
        Encoding.UTF8.GetByteCount(CreateSlideMasterXml()) +
        Encoding.UTF8.GetByteCount(CreateThemeXml());

    private static string CreateContentTypes(string mainContentType) =>
        $"""
         <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
         <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
           <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
           <Default Extension="xml" ContentType="application/xml"/>
           <Default Extension="png" ContentType="image/png"/>
           <Override PartName="/ppt/presentation.xml" ContentType="{mainContentType}"/>
           <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
         </Types>
         """;

    private static string CreatePackageRelationships(bool useStrictNamespaces = false)
    {
        var relationshipType = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                   <Relationship Id="rId1" Type="{{relationshipType}}" Target="ppt/presentation.xml"/>
                 </Relationships>
                 """;
    }

    private static string CreatePresentationXml(bool useStrictNamespaces = false)
    {
        var presentationNamespace = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/presentationml/main"
            : "http://schemas.openxmlformats.org/presentationml/2006/main";
        var drawingNamespace = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/drawingml/main"
            : "http://schemas.openxmlformats.org/drawingml/2006/main";
        var relationshipNamespace = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <p:presentation xmlns:a="{{drawingNamespace}}" xmlns:r="{{relationshipNamespace}}" xmlns:p="{{presentationNamespace}}">
                   <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                   <p:sldSz cx="12192000" cy="6858000"/>
                   <p:notesSz cx="6858000" cy="9144000"/>
                 </p:presentation>
                 """;
    }

    private static string CreatePresentationRelationships(bool useStrictNamespaces = false)
    {
        var relationshipType = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships/slide"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                   <Relationship Id="rId1" Type="{{relationshipType}}" Target="slides/slide1.xml"/>
                 </Relationships>
                 """;
    }

    private static string CreateSlideXml(IReadOnlyList<int> picturePartIndexes)
    {
        var pictures = new StringBuilder();
        for (var index = 0; index < picturePartIndexes.Count; index++)
        {
            pictures.Append(CultureInvariantPictureXml(index + 2, picturePartIndexes[index] + 1));
        }

        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                   <p:cSld><p:spTree>
                     <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                     <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                     {{pictures}}
                   </p:spTree></p:cSld>
                   <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
                 </p:sld>
                 """;
    }

    private static string CreateTextSlideXml(
        string text,
        string? documentTypeDeclaration = null,
        bool useStrictNamespaces = false)
    {
        var presentationNamespace = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/presentationml/main"
            : "http://schemas.openxmlformats.org/presentationml/2006/main";
        var drawingNamespace = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/drawingml/main"
            : "http://schemas.openxmlformats.org/drawingml/2006/main";
        var relationshipNamespace = useStrictNamespaces
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 {{documentTypeDeclaration}}
                 <p:sld xmlns:a="{{drawingNamespace}}" xmlns:r="{{relationshipNamespace}}" xmlns:p="{{presentationNamespace}}">
                   <p:cSld><p:spTree>
                     <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                     <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                     <p:sp>
                       <p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                       <p:spPr/>
                       <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>{{text}}</a:t></a:r></a:p></p:txBody>
                     </p:sp>
                   </p:spTree></p:cSld>
                   <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
                 </p:sld>
                 """;
    }

    private static string CreateSharedPartContentTypes() =>
        $"""
         <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
         <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
           <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
           <Default Extension="xml" ContentType="application/xml"/>
           <Override PartName="/ppt/presentation.xml" ContentType="{PresentationContentType}"/>
           <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
           <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
           <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
           <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
           <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
         </Types>
         """;

    private static string CreateTwoSlidePresentationXml() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1"/>
            <p:sldId id="257" r:id="rId2"/>
          </p:sldIdLst>
          <p:sldSz cx="12192000" cy="6858000"/>
          <p:notesSz cx="6858000" cy="9144000"/>
        </p:presentation>
        """;

    private static string CreateTwoSlidePresentationRelationships() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
        </Relationships>
        """;

    private static string CreateInheritedTextSlideXml(int slideNumber) =>
        $$"""
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
              {{CreatePlaceholderTextShapeXml(2, "body", "Body", slideNumber)}}
              {{CreatePlaceholderTextShapeXml(3, "title", "Title", slideNumber)}}
            </p:spTree></p:cSld>
            <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
          </p:sld>
          """;

    private static string CreatePlaceholderTextShapeXml(
        int shapeId,
        string placeholderType,
        string text,
        int slideNumber)
    {
        var placeholderIndex = placeholderType == "title" ? 1 : 2;
        return $$"""
                 <p:sp>
                   <p:nvSpPr><p:cNvPr id="{{shapeId}}" name="{{text}}"/><p:cNvSpPr/><p:nvPr><p:ph type="{{placeholderType}}" idx="{{placeholderIndex}}"/></p:nvPr></p:nvSpPr>
                   <p:spPr/>
                   <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>{{text}} {{slideNumber}}</a:t></a:r></a:p></p:txBody>
                 </p:sp>
                 """;
    }

    private static string CreateSlideLayoutXml() =>
        $$"""
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
              {{CreateInheritedPlaceholderXml(2, "title", 1, 100)}}
            </p:spTree></p:cSld>
          </p:sldLayout>
          """;

    private static string CreateSlideMasterXml() =>
        $$"""
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
            <p:cSld><p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
              {{CreateInheritedPlaceholderXml(2, "body", 2, 1_000)}}
            </p:spTree></p:cSld>
          </p:sldMaster>
          """;

    private static string CreateInheritedPlaceholderXml(
        int shapeId,
        string placeholderType,
        int placeholderIndex,
        int y) =>
        $$"""
          <p:sp>
            <p:nvSpPr><p:cNvPr id="{{shapeId}}" name="{{placeholderType}}"/><p:cNvSpPr/><p:nvPr><p:ph type="{{placeholderType}}" idx="{{placeholderIndex}}"/></p:nvPr></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="100" y="{{y}}"/><a:ext cx="1000" cy="100"/></a:xfrm></p:spPr>
          </p:sp>
          """;

    private static string CreateSlideLayoutRelationship() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
        </Relationships>
        """;

    private static string CreateSlideMasterRelationship() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
        </Relationships>
        """;

    private static string CreateThemeRelationship() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
        </Relationships>
        """;

    private static string CreateThemeXml() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
          <a:themeElements>
            <a:clrScheme name="Test Colors">
              <a:dk1><a:srgbClr val="000000"/></a:dk1>
              <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
              <a:accent1><a:srgbClr val="1A2B3C"/></a:accent1>
            </a:clrScheme>
          </a:themeElements>
        </a:theme>
        """;

    private static string CultureInvariantPictureXml(int shapeId, int imagePartIndex) =>
        $"""
         <p:pic>
           <p:nvPicPr><p:cNvPr id="{shapeId}" name="Image {shapeId}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
           <p:blipFill><a:blip r:embed="rIdImage{imagePartIndex}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
           <p:spPr><a:xfrm><a:off x="{shapeId * 100}" y="{shapeId * 100}"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
         </p:pic>
         """;

    private static string CreateSlideRelationships(int imagePartCount)
    {
        var relationships = new StringBuilder();
        for (var index = 1; index <= imagePartCount; index++)
        {
            relationships.Append(
                CultureInfo.InvariantCulture,
                $"<Relationship Id=\"rIdImage{index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image{index}.png\"/>");
        }

        return $$"""
                 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{{relationships}}</Relationships>
                 """;
    }

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

    [SuppressMessage(
        "ReSharper",
        "MemberCanBePrivate.Global",
        Justification = "MSTest requires a public TestContext property for runtime injection.")]
    [SuppressMessage(
        "ReSharper",
        "AutoPropertyCanBeMadeGetOnly.Global",
        Justification = "MSTest sets TestContext through this property at runtime.")]
    public TestContext TestContext { get; set; } = null!;
}
