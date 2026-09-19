using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using DrawingColor = System.Drawing.Color;
using DrawingImage = System.Drawing.Image;
using DrawingImageFormat = System.Drawing.Imaging.ImageFormat;

namespace Zen.WinUI;

internal readonly record struct CropArea(int X, int Y, int Width, int Height);

internal static class ImageProcessor
{
    internal static void Process(string operation, string input, string output, CropArea cropArea, int radius)
    {
        using var source = DrawingImage.FromFile(input);
        switch (operation)
        {
            case "image_jpg": SaveJpeg(source, output); break;
            case "png_pdf":
            case "jpg_pdf": ImagePdfWriter.Write(source, output); break;
            case "crop": Crop(source, output, cropArea); break;
            case "rounded": WriteMasked(source, output, false, radius); break;
            case "circle": WriteMasked(source, output, true, radius); break;
            default: throw new InvalidOperationException("Cette opération n’est pas prise en charge.");
        }
    }

    private static void SaveJpeg(DrawingImage source, string output)
    {
        using var result = new Bitmap(source.Width, source.Height, PixelFormat.Format24bppRgb);
        using (var graphics = Graphics.FromImage(result))
        {
            graphics.Clear(DrawingColor.White);
            graphics.DrawImage(source, 0, 0, source.Width, source.Height);
        }
        result.Save(output, DrawingImageFormat.Jpeg);
    }

    private static void Crop(DrawingImage source, string output, CropArea requestedArea)
    {
        var area = new Rectangle(
            requestedArea.X,
            requestedArea.Y,
            Math.Min(requestedArea.Width, source.Width - requestedArea.X),
            Math.Min(requestedArea.Height, source.Height - requestedArea.Y));
        if (area.Width <= 0 || area.Height <= 0)
            throw new InvalidOperationException("Les dimensions de rognage doivent rester dans l’image.");

        using var result = new Bitmap(area.Width, area.Height);
        using (var graphics = Graphics.FromImage(result))
        {
            graphics.DrawImage(
                source,
                new Rectangle(0, 0, result.Width, result.Height),
                area,
                GraphicsUnit.Pixel);
        }
        result.Save(output, DrawingImageFormat.Png);
    }

    private static void WriteMasked(DrawingImage source, string output, bool circle, int radius)
    {
        var size = circle ? Math.Min(source.Width, source.Height) : 0;
        var width = circle ? size : source.Width;
        var height = circle ? size : source.Height;
        using var result = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(result);
        using var path = CreateMaskPath(width, height, circle, radius);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.SetClip(path);
        var sourceArea = circle
            ? new Rectangle((source.Width - size) / 2, (source.Height - size) / 2, size, size)
            : new Rectangle(0, 0, source.Width, source.Height);
        graphics.DrawImage(source, new Rectangle(0, 0, width, height), sourceArea, GraphicsUnit.Pixel);
        result.Save(output, DrawingImageFormat.Png);
    }

    private static GraphicsPath CreateMaskPath(int width, int height, bool circle, int requestedRadius)
    {
        var path = new GraphicsPath();
        if (circle)
        {
            path.AddEllipse(0, 0, width - 1, height - 1);
            return path;
        }

        var radius = Math.Min(requestedRadius, Math.Min(width, height) / 2);
        var diameter = radius * 2;
        path.AddArc(0, 0, diameter, diameter, 180, 90);
        path.AddArc(width - diameter, 0, diameter, diameter, 270, 90);
        path.AddArc(width - diameter, height - diameter, diameter, diameter, 0, 90);
        path.AddArc(0, height - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}

internal static class ImagePdfWriter
{
    internal static void Write(DrawingImage image, string output)
    {
        var jpeg = CreateJpeg(image);
        var bytes = new List<byte>();
        var offsets = new List<long>();
        void WriteText(string value) => bytes.AddRange(Encoding.ASCII.GetBytes(value));
        void WriteObject(int id, string body)
        {
            offsets.Add(bytes.Count);
            WriteText($"{id} 0 obj\n{body}\nendobj\n");
        }

        WriteText("%PDF-1.4\n");
        WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
        WriteObject(3, $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {image.Width} {image.Height}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>");
        offsets.Add(bytes.Count);
        WriteText($"4 0 obj\n<< /Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {jpeg.Length} >>\nstream\n");
        bytes.AddRange(jpeg);
        WriteText("\nendstream\nendobj\n");
        var content = $"q\n{image.Width} 0 0 {image.Height} 0 0 cm\n/Im0 Do\nQ\n";
        WriteObject(5, $"<< /Length {Encoding.ASCII.GetByteCount(content)} >>\nstream\n{content}endstream");
        WriteCrossReference(bytes, offsets, WriteText);
        File.WriteAllBytes(output, bytes.ToArray());
    }

    private static byte[] CreateJpeg(DrawingImage image)
    {
        using var stream = new MemoryStream();
        using var whiteBackground = new Bitmap(image.Width, image.Height);
        using (var graphics = Graphics.FromImage(whiteBackground))
        {
            graphics.Clear(DrawingColor.White);
            graphics.DrawImage(image, 0, 0, image.Width, image.Height);
        }
        whiteBackground.Save(stream, DrawingImageFormat.Jpeg);
        return stream.ToArray();
    }

    private static void WriteCrossReference(List<byte> bytes, IReadOnlyList<long> offsets, Action<string> write)
    {
        var crossReferenceOffset = bytes.Count;
        write("xref\n0 6\n0000000000 65535 f \n");
        foreach (var offset in offsets) write($"{offset:D10} 00000 n \n");
        write($"trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{crossReferenceOffset}\n%%EOF");
    }
}
