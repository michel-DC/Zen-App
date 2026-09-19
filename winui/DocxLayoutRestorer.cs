using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Zen.WinUI;

internal static class DocxLayoutRestorer
{
    private static readonly Regex EmailPattern = new(
        @"^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex WebAddressPattern = new(
        @"^(?:https?://|www\.)\S+$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    internal static void Restore(string pdfPath, string docxPath)
    {
        var sourcePages = PdfTextProcessor.ReadPages(pdfPath);
        var sourceLines = sourcePages
            .SelectMany(page => page.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();
        if (sourceLines.Count == 0) return;

        using var package = WordprocessingDocument.Open(docxPath, true);
        var mainPart = package.MainDocumentPart
            ?? throw new InvalidOperationException("Le document Word converti ne contient pas de partie principale.");
        var wordDocument = mainPart.Document
            ?? throw new InvalidOperationException("Le document Word converti ne contient pas de contenu lisible.");
        var body = wordDocument.Body;
        if (body is null) return;

        RestoreLineBreaks(body, sourceLines);
        DocxParallelSectionRestorer.Restore(body);
        DocxPageRestorer.Restore(body, sourcePages);
        RestoreHyperlinks(mainPart, body);
        wordDocument.Save();
    }

    private static void RestoreLineBreaks(Body body, IReadOnlyList<string> sourceLines)
    {
        var lineCursor = 0;
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var paragraphKey = Normalize(paragraph.InnerText);
            if (paragraphKey.Length == 0) continue;
            var match = FindSourceLines(sourceLines, lineCursor, paragraphKey);
            if (match is null) continue;
            lineCursor = match.Value.End + 1;
            if (match.Value.End <= match.Value.Start) continue;
            InsertBreaks(paragraph, CreateBoundaries(sourceLines, match.Value));
        }
    }

    private static IReadOnlyList<int> CreateBoundaries(IReadOnlyList<string> lines, (int Start, int End) match)
    {
        var boundaries = new List<int>();
        var characterCount = 0;
        for (var index = match.Start; index < match.End; index++)
        {
            characterCount += Normalize(lines[index]).Length;
            boundaries.Add(characterCount);
        }
        return boundaries;
    }

    private static (int Start, int End)? FindSourceLines(
        IReadOnlyList<string> lines,
        int cursor,
        string paragraphKey)
    {
        for (var start = cursor; start < Math.Min(lines.Count, cursor + 4); start++)
        {
            var candidate = new StringBuilder();
            for (var end = start; end < lines.Count && end < start + 40; end++)
            {
                candidate.Append(Normalize(lines[end]));
                var candidateText = candidate.ToString();
                if (candidateText == paragraphKey) return (start, end);
                if (!paragraphKey.StartsWith(candidateText, StringComparison.Ordinal)) break;
            }
        }
        return null;
    }

    private static void InsertBreaks(Paragraph paragraph, IReadOnlyCollection<int> boundaries)
    {
        var pending = new HashSet<int>(boundaries);
        var seenCharacters = 0;
        foreach (var textNode in paragraph.Descendants<Text>().ToList())
        {
            var splitOffsets = FindSplitOffsets(textNode.Text, pending, ref seenCharacters);
            if (splitOffsets.Count > 0) SplitTextNode(textNode, splitOffsets);
        }
    }

    private static List<int> FindSplitOffsets(string value, ISet<int> boundaries, ref int seenCharacters)
    {
        var offsets = new List<int>();
        for (var index = 0; index < value.Length; index++)
        {
            if (char.IsWhiteSpace(value[index])) continue;
            seenCharacters++;
            if (boundaries.Remove(seenCharacters)) offsets.Add(index + 1);
        }
        return offsets;
    }

    private static void SplitTextNode(Text textNode, IEnumerable<int> splitOffsets)
    {
        var value = textNode.Text;
        var start = 0;
        foreach (var offset in splitOffsets)
        {
            InsertTextBefore(textNode, value[start..offset].TrimEnd());
            textNode.InsertBeforeSelf(new Break());
            start = offset;
            while (start < value.Length && char.IsWhiteSpace(value[start])) start++;
        }
        InsertTextBefore(textNode, value[start..]);
        textNode.Remove();
    }

    private static void InsertTextBefore(Text anchor, string value)
    {
        if (value.Length > 0)
            anchor.InsertBeforeSelf(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static void RestoreHyperlinks(MainDocumentPart mainPart, Body body)
    {
        foreach (var run in body.Descendants<Run>().ToList())
        {
            if (run.Ancestors<Hyperlink>().Any()) continue;
            var target = DetectLinkTarget(run.InnerText.Trim());
            if (target is null) continue;
            var relationship = mainPart.AddHyperlinkRelationship(target, true);
            run.InsertAfterSelf(new Hyperlink(run.CloneNode(true)) { Id = relationship.Id });
            run.Remove();
        }
    }

    private static Uri? DetectLinkTarget(string label)
    {
        if (EmailPattern.IsMatch(label)) return new Uri($"mailto:{label}");
        if (!WebAddressPattern.IsMatch(label)) return null;
        return new Uri(label.StartsWith("www.", StringComparison.OrdinalIgnoreCase) ? $"https://{label}" : label);
    }

    private static string Normalize(string value) =>
        string.Concat(value.Where(character => !char.IsWhiteSpace(character)));
}
