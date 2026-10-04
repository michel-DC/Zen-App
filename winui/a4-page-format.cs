using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Zen.WinUI;

internal static class A4PageFormat
{
    internal const double WidthInPoints = 210 * 72 / 25.4;
    internal const double HeightInPoints = 297 * 72 / 25.4;
    private static readonly uint WidthInTwips = (uint)Math.Round(WidthInPoints * 20);
    private static readonly uint HeightInTwips = (uint)Math.Round(HeightInPoints * 20);

    internal static void PrepareWordCopy(string input, string output)
    {
        File.Copy(input, output);
        using var document = WordprocessingDocument.Open(output, true);
        var mainDocument = document.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("Le document DOCX ne contient pas de contenu lisible.");
        var body = mainDocument.Body
            ?? throw new InvalidOperationException("Le document DOCX ne contient pas de contenu lisible.");
        if (body.GetFirstChild<SectionProperties>() is null)
            body.AppendChild(new SectionProperties());

        foreach (var section in body.Descendants<SectionProperties>())
        {
            var pageSize = section.GetFirstChild<PageSize>();
            if (pageSize is null)
            {
                pageSize = new PageSize();
                section.AddChild(pageSize, true);
            }

            var landscape = pageSize.Orient?.Value == PageOrientationValues.Landscape
                || (pageSize.Orient is null && pageSize.Width?.Value > pageSize.Height?.Value);
            pageSize.Width = landscape ? HeightInTwips : WidthInTwips;
            pageSize.Height = landscape ? WidthInTwips : HeightInTwips;
            pageSize.Orient = landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
            pageSize.Code = null;
        }

        mainDocument.Save();
    }
}
