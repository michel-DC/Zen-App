using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Zen.WinUI;

internal static class DocxPageRestorer
{
    internal static void Restore(Body body, IReadOnlyList<string> sourcePages)
    {
        if (sourcePages.Count < 2) return;

        var paragraphs = body.Elements<Paragraph>().ToList();
        var pageKeys = sourcePages.Select(Normalize).ToList();
        var previousBoundary = 0;
        var unresolvedPages = new List<int>();

        for (var pageIndex = 1; pageIndex < sourcePages.Count; pageIndex++)
        {
            var boundary = FindBoundary(paragraphs, previousBoundary, pageIndex, sourcePages, pageKeys);
            if (boundary is null)
            {
                unresolvedPages.Add(pageIndex);
                continue;
            }

            if (!HasPageBreakImmediatelyBefore(paragraphs, boundary.Value.ParagraphIndex))
                InsertPageBoundary(paragraphs[boundary.Value.ParagraphIndex], boundary.Value.CharacterOffset);
            previousBoundary = boundary.Value.ParagraphIndex;
        }

        RestoreTrailingVisualPage(paragraphs, sourcePages, unresolvedPages, previousBoundary);
    }

    private static PageBoundary? FindBoundary(
        IReadOnlyList<Paragraph> paragraphs,
        int paragraphStart,
        int pageIndex,
        IReadOnlyList<string> sourcePages,
        IReadOnlyList<string> pageKeys)
    {
        foreach (var line in SplitLines(sourcePages[pageIndex]).Take(3))
        {
            var lineKey = Normalize(line);
            if (lineKey.Length < 10 || !IsUniqueToPage(lineKey, pageIndex, pageKeys)) continue;

            for (var index = paragraphStart; index < paragraphs.Count; index++)
            {
                var paragraphKey = Normalize(GetFlowText(paragraphs[index]));
                if (paragraphKey.Length < 6) continue;
                var offset = paragraphKey.IndexOf(lineKey, StringComparison.Ordinal);
                if (offset >= 0) return new PageBoundary(index, offset);
                if (lineKey.StartsWith(paragraphKey, StringComparison.Ordinal))
                    return new PageBoundary(index, 0);
            }
        }

        return null;
    }

    private static bool IsUniqueToPage(string lineKey, int pageIndex, IReadOnlyList<string> pageKeys) =>
        pageKeys.Where((_, index) => index != pageIndex)
            .All(page => !page.Contains(lineKey, StringComparison.Ordinal));

    private static void RestoreTrailingVisualPage(
        IReadOnlyList<Paragraph> paragraphs,
        IReadOnlyList<string> sourcePages,
        IReadOnlyCollection<int> unresolvedPages,
        int previousBoundary)
    {
        var lastPage = sourcePages.Count - 1;
        if (!unresolvedPages.Contains(lastPage) || Normalize(sourcePages[lastPage]).Length > 80) return;

        for (var index = paragraphs.Count - 1; index > previousBoundary; index--)
        {
            if (!paragraphs[index].Descendants<Drawing>().Any()) continue;
            AddPageBreakBefore(paragraphs[index]);
            return;
        }
    }

    private static void InsertPageBoundary(Paragraph paragraph, int characterOffset)
    {
        if (characterOffset == 0)
        {
            AddPageBreakBefore(paragraph);
            return;
        }

        var seenCharacters = 0;
        foreach (var textNode in FlowTextNodes(paragraph).ToList())
        {
            for (var index = 0; index < textNode.Text.Length; index++)
            {
                if (char.IsWhiteSpace(textNode.Text[index])) continue;
                if (seenCharacters++ != characterOffset) continue;
                SplitTextNodeWithPageBreak(textNode, index);
                return;
            }
        }
    }

    private static void SplitTextNodeWithPageBreak(Text textNode, int offset)
    {
        var value = textNode.Text;
        if (offset > 0)
            textNode.InsertBeforeSelf(new Text(value[..offset]) { Space = SpaceProcessingModeValues.Preserve });
        textNode.InsertBeforeSelf(new Break { Type = BreakValues.Page });
        textNode.InsertBeforeSelf(new Text(value[offset..]) { Space = SpaceProcessingModeValues.Preserve });
        textNode.Remove();
    }

    private static void AddPageBreakBefore(Paragraph paragraph)
    {
        var properties = paragraph.GetFirstChild<ParagraphProperties>();
        if (properties is null)
        {
            properties = new ParagraphProperties();
            paragraph.PrependChild(properties);
        }
        properties.PageBreakBefore ??= new PageBreakBefore();
    }

    private static bool HasPageBreakImmediatelyBefore(
        IReadOnlyList<Paragraph> paragraphs,
        int paragraphIndex)
    {
        for (var index = paragraphIndex - 1; index >= Math.Max(0, paragraphIndex - 3); index--)
        {
            if (paragraphs[index].Descendants<Break>()
                .Any(pageBreak => pageBreak.Type?.Value == BreakValues.Page))
                return true;
            if (Normalize(GetFlowText(paragraphs[index])).Length > 0) return false;
        }
        return false;
    }

    private static string GetFlowText(Paragraph paragraph) =>
        string.Concat(FlowTextNodes(paragraph).Select(text => text.Text));

    private static IEnumerable<Text> FlowTextNodes(Paragraph paragraph) =>
        paragraph.Descendants<Text>().Where(text => !text.Ancestors<TextBoxContent>().Any());

    private static IEnumerable<string> SplitLines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n')
            .Where(line => !string.IsNullOrWhiteSpace(line));

    private static string Normalize(string value) =>
        string.Concat(value.Where(character => !char.IsWhiteSpace(character)));

    private readonly record struct PageBoundary(int ParagraphIndex, int CharacterOffset);
}
