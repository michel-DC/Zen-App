using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Zen.WinUI;

internal static class DocxParallelSectionRestorer
{
    private const int ParallelIndentThreshold = 700;

    internal static void Restore(Body body)
    {
        var paragraphs = body.Elements<Paragraph>().ToList();
        var previousSectionEnd = -1;

        for (var sectionEnd = 0; sectionEnd < paragraphs.Count; sectionEnd++)
        {
            var sectionProperties = paragraphs[sectionEnd].ParagraphProperties?.SectionProperties;
            if (sectionProperties is null) continue;

            var columns = sectionProperties.GetFirstChild<Columns>();
            var columnCount = columns?.ColumnCount?.Value ?? 1;
            if (columnCount > 1)
                RestoreSection(body, paragraphs, previousSectionEnd + 1, sectionEnd, sectionProperties);
            previousSectionEnd = sectionEnd;
        }
    }

    private static void RestoreSection(
        Body body,
        IReadOnlyList<Paragraph> paragraphs,
        int start,
        int end,
        SectionProperties sectionProperties)
    {
        var split = FindRightColumnStart(paragraphs, start, end);
        if (split is null) return;

        var left = MeaningfulParagraphs(paragraphs, start, split.Value)
            .Where(paragraph => !IsImportArtifact(paragraph));
        var right = MeaningfulParagraphs(paragraphs, split.Value, end)
            .Where(paragraph => !IsImportArtifact(paragraph));
        var table = CreateParallelTable(left, right);

        paragraphs[start].InsertBeforeSelf(table);
        for (var index = start; index < end; index++)
            paragraphs[index].Remove();

        var columns = sectionProperties.GetFirstChild<Columns>();
        if (columns is not null) columns.ColumnCount = 1;
    }

    private static int? FindRightColumnStart(
        IReadOnlyList<Paragraph> paragraphs,
        int start,
        int end)
    {
        var midpoint = start + ((end - start) / 3);
        for (var index = midpoint; index < end; index++)
        {
            if (Normalize(paragraphs[index].InnerText).Length == 0) continue;
            var indentation = paragraphs[index].ParagraphProperties?.Indentation?.Left?.Value;
            if (int.TryParse(indentation, out var left) && left >= ParallelIndentThreshold)
                return index;
        }
        return null;
    }

    private static IEnumerable<Paragraph> MeaningfulParagraphs(
        IReadOnlyList<Paragraph> paragraphs,
        int start,
        int end)
    {
        var selected = paragraphs.Skip(start).Take(end - start).ToList();
        while (selected.Count > 0 && Normalize(selected[0].InnerText).Length == 0) selected.RemoveAt(0);
        while (selected.Count > 0 && Normalize(selected[^1].InnerText).Length == 0) selected.RemoveAt(selected.Count - 1);
        return selected;
    }

    private static bool IsImportArtifact(Paragraph paragraph) =>
        string.Equals(paragraph.InnerText.Trim(), "pour", StringComparison.OrdinalIgnoreCase);

    private static Table CreateParallelTable(
        IEnumerable<Paragraph> leftParagraphs,
        IEnumerable<Paragraph> rightParagraphs)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "8950", Type = TableWidthUnitValues.Dxa },
                new TableLayout { Type = TableLayoutValues.Fixed },
                CreateBorders()),
            new TableGrid(new GridColumn { Width = "5000" }, new GridColumn { Width = "3950" }));

        var row = new TableRow(new TableRowProperties(new CantSplit()));
        row.Append(CreateCell(leftParagraphs, "5000", true));
        row.Append(CreateCell(rightParagraphs, "3950", false));
        table.Append(row);
        return table;
    }

    private static TableCell CreateCell(
        IEnumerable<Paragraph> paragraphs,
        string width,
        bool hasDivider)
    {
        var borders = new TableCellBorders();
        if (hasDivider)
            borders.RightBorder = new RightBorder { Val = BorderValues.Single, Size = 8, Color = "000000" };

        var cell = new TableCell(
            new TableCellProperties(
                new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa },
                new TableCellMargin(
                    new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new StartMargin { Width = hasDivider ? "0" : "120", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new EndMargin { Width = "120", Type = TableWidthUnitValues.Dxa }),
                borders));

        foreach (var paragraph in paragraphs)
        {
            var clone = (Paragraph)paragraph.CloneNode(true);
            clone.ParagraphProperties?.RemoveAllChildren<SectionProperties>();
            foreach (var drawing in clone.Descendants<AlternateContent>()
                .Where(content => content.Descendants<Drawing>().Any()).ToList())
                drawing.Remove();
            ResetColumnIndent(clone);
            cell.Append(clone);
        }
        if (!cell.Elements<Paragraph>().Any()) cell.Append(new Paragraph());
        return cell;
    }

    private static TableBorders CreateBorders() => new(
        new TopBorder { Val = BorderValues.Nil },
        new LeftBorder { Val = BorderValues.Nil },
        new BottomBorder { Val = BorderValues.Nil },
        new RightBorder { Val = BorderValues.Nil },
        new InsideHorizontalBorder { Val = BorderValues.Nil },
        new InsideVerticalBorder { Val = BorderValues.Nil });

    private static void ResetColumnIndent(Paragraph paragraph)
    {
        var indentation = paragraph.ParagraphProperties?.Indentation;
        if (indentation is null) return;
        indentation.Left = "0";
        indentation.Right = "0";
    }

    private static string Normalize(string value) =>
        string.Concat(value.Where(character => !char.IsWhiteSpace(character)));
}
