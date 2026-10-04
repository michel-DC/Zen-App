using System;
using System.Globalization;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Zen.WinUI;

internal static class PdfA4PageNormalizer
{
    internal static void Normalize(PdfPage page)
    {
        var media = page.MediaBox;
        var crop = page.Elements.ContainsKey("/CropBox") ? page.CropBox : media;
        var left = Math.Max(crop.X1, media.X1);
        var bottom = Math.Max(crop.Y1, media.Y1);
        var width = Math.Min(crop.X2, media.X2) - left;
        var height = Math.Min(crop.Y2, media.Y2) - bottom;
        if (!double.IsFinite(width) || !double.IsFinite(height) || width <= 0 || height <= 0)
            throw new InvalidOperationException("Une page PDF contient des dimensions invalides.");

        var landscape = width > height;
        var pageWidth = landscape ? A4PageFormat.HeightInPoints : A4PageFormat.WidthInPoints;
        var pageHeight = landscape ? A4PageFormat.WidthInPoints : A4PageFormat.HeightInPoints;
        var scale = Math.Min(pageWidth / width, pageHeight / height);
        var x = (pageWidth - width * scale) / 2 - left * scale;
        var y = (pageHeight - height * scale) / 2 - bottom * scale;
        var prefix = string.Create(CultureInfo.InvariantCulture,
            $"q\n{scale:0.##########} 0 0 {scale:0.##########} {x:0.##########} {y:0.##########} cm\n{left:0.##########} {bottom:0.##########} {width:0.##########} {height:0.##########} re W n\n");
        page.Contents.PrependContent().CreateStream(Encoding.ASCII.GetBytes(prefix));
        page.Contents.AppendContent().CreateStream(Encoding.ASCII.GetBytes("\nQ\n"));

        for (var annotationIndex = 0; annotationIndex < page.Annotations.Count; annotationIndex++)
        {
            var annotation = page.Annotations[annotationIndex];
            var rectangle = annotation.Rectangle;
            annotation.Rectangle = new PdfRectangle(
                new XPoint(rectangle.X1 * scale + x, rectangle.Y1 * scale + y),
                new XPoint(rectangle.X2 * scale + x, rectangle.Y2 * scale + y));
            var points = annotation.Elements.GetArray("/QuadPoints");
            if (points is null) continue;
            for (var index = 0; index < points.Elements.Count; index++)
                points.Elements[index] = new PdfReal(points.Elements.GetReal(index) * scale + (index % 2 == 0 ? x : y));
        }

        var a4 = new PdfRectangle(new XRect(0, 0, pageWidth, pageHeight));
        page.MediaBox = a4;
        page.CropBox = a4;
        foreach (var box in new[] { "/BleedBox", "/TrimBox", "/ArtBox" })
            if (page.Elements.ContainsKey(box)) page.Elements.SetRectangle(box, a4);
        page.Elements.Remove("/UserUnit");
    }
}
