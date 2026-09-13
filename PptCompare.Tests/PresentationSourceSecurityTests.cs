using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
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
        bool includeSlidePart = true)
    {
        imageParts ??= [];
        picturePartIndexes ??= [];
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var package = new ZipArchive(stream, ZipArchiveMode.Create);

        WriteTextEntry(
            package,
            "[Content_Types].xml",
            CreateContentTypes(mainContentType));
        WriteTextEntry(package, "_rels/.rels", CreatePackageRelationships());
        if (!includePresentationPart)
        {
            return;
        }

        WriteTextEntry(package, "ppt/presentation.xml", CreatePresentationXml());
        WriteTextEntry(
            package,
            "ppt/_rels/presentation.xml.rels",
            CreatePresentationRelationships());
        if (!includeSlidePart)
        {
            return;
        }

        WriteTextEntry(
            package,
            "ppt/slides/slide1.xml",
            CreateSlideXml(picturePartIndexes));
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
