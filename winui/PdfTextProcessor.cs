using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Zen.WinUI;

internal static class PdfTextProcessor
{
    internal static string ReadText(string input) =>
        string.Join($"{Environment.NewLine}{Environment.NewLine}", ReadPages(input));

    internal static IReadOnlyList<string> ReadPages(string input)
    {
        using var document = PdfDocument.Open(input);
        var pages = document.GetPages()
            .Select(page => ContentOrderTextExtractor.GetText(page))
            .ToList();
        if (pages.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException(
                "Ce PDF ne contient pas de texte sélectionnable. Utilisez l’extraction depuis une image pour un document numérisé.");
        return pages;
    }
}
