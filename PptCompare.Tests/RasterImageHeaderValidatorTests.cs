using System.Buffers.Binary;
using PptCompare.Controls;

namespace PptCompare.Tests;

[TestClass]
public sealed class RasterImageHeaderValidatorTests
{
    [TestMethod]
    public void ValidPngHeaderReturnsDimensions()
    {
        var bytes = CreatePngHeader(1_600, 900);

        var isValid = RasterImageHeaderValidator.TryValidate(bytes, out var header);

        Assert.IsTrue(isValid);
        Assert.AreEqual(RasterImageFormat.Png, header.Format);
        Assert.AreEqual(1_600, header.PixelWidth);
        Assert.AreEqual(900, header.PixelHeight);
    }

    [TestMethod]
    public void ValidJpegHeaderReturnsDimensions()
    {
        var bytes = CreateJpegHeader(1_920, 1_080);

        var isValid = RasterImageHeaderValidator.TryValidate(bytes, out var header);

        Assert.IsTrue(isValid);
        Assert.AreEqual(RasterImageFormat.Jpeg, header.Format);
        Assert.AreEqual(1_920, header.PixelWidth);
        Assert.AreEqual(1_080, header.PixelHeight);
    }

    [TestMethod]
    public void TruncatedOrMalformedHeadersAreRejected()
    {
        var truncatedPng = CreatePngHeader(100, 100)[..20];
        byte[] truncatedJpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x20, 0x00];
        var malformedPng = CreatePngHeader(100, 100);
        malformedPng[11] = 12;

        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(truncatedPng, out _));
        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(truncatedJpeg, out _));
        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(malformedPng, out _));
    }

    [TestMethod]
    public void UnsupportedRasterFormatIsRejected()
    {
        byte[] gifHeader = "GIF89a"u8.ToArray();

        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(gifHeader, out _));
    }

    [TestMethod]
    public void ZeroAndOutOfRangeDimensionsAreRejected()
    {
        var zeroWidth = CreatePngHeader(0, 100);
        var signedNegativeRepresentation = CreatePngHeader(uint.MaxValue, 100);
        var excessiveDimension = CreatePngHeader(
            RasterImageHeaderValidator.MaxSourceDimension + 1U,
            1);

        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(zeroWidth, out _));
        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(
            signedNegativeRepresentation,
            out _));
        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(excessiveDimension, out _));
    }

    [TestMethod]
    public void ExcessivePixelCountIsRejected()
    {
        var bytes = CreatePngHeader(8_000, 8_000);

        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(bytes, out _));
    }

    [TestMethod]
    public void ExtremeAspectRatioIsRejected()
    {
        var bytes = CreatePngHeader(
            RasterImageHeaderValidator.MaxAspectRatio + 1U,
            1);

        Assert.IsFalse(RasterImageHeaderValidator.TryValidate(bytes, out _));
    }

    [TestMethod]
    public void DecodePixelEstimateCapsTheLongestSideWithoutUpscaling()
    {
        var largeLandscape = new RasterImageHeader(RasterImageFormat.Png, 4_000, 2_000);
        var largePortrait = new RasterImageHeader(RasterImageFormat.Jpeg, 2_000, 4_000);
        var smallImage = new RasterImageHeader(RasterImageFormat.Png, 800, 600);

        Assert.AreEqual(
            1_600L * 800,
            RasterImageHeaderValidator.CalculateDecodedPixelCount(largeLandscape));
        Assert.AreEqual(
            1_600L * 800,
            RasterImageHeaderValidator.CalculateDecodedPixelCount(largePortrait));
        Assert.AreEqual(
            800L * 600,
            RasterImageHeaderValidator.CalculateDecodedPixelCount(smallImage));
    }

    [TestMethod]
    public void AggregateDecodedPixelBudgetRejectsExcessImages()
    {
        long retainedPixels = 0;
        const long decodedPixelsPerImage = 1_600L * 1_000;

        for (var index = 0; index < 12; index++)
        {
            Assert.IsTrue(RasterImageHeaderValidator.TryReserveDecodedPixels(
                ref retainedPixels,
                decodedPixelsPerImage));
        }

        Assert.AreEqual(19_200_000, retainedPixels);
        Assert.IsFalse(RasterImageHeaderValidator.TryReserveDecodedPixels(
            ref retainedPixels,
            decodedPixelsPerImage));
        Assert.AreEqual(19_200_000, retainedPixels);
    }

    private static byte[] CreatePngHeader(uint width, uint height)
    {
        byte[] bytes =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        ];
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(16, 4), width);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(20, 4), height);
        return bytes;
    }

    private static byte[] CreateJpegHeader(ushort width, ushort height)
    {
        byte[] bytes =
        [
            0xFF, 0xD8,
            0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00,
            0xFF, 0xC0, 0x00, 0x0B, 0x08,
            0x00, 0x00,
            0x00, 0x00,
            0x01, 0x01, 0x11, 0x00
        ];
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(13, 2), height);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(15, 2), width);
        return bytes;
    }
}
