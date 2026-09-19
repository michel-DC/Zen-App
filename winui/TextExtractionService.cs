using System;
using System.Threading.Tasks;

namespace Zen.WinUI;

internal static class TextExtractionService
{
    internal static Task<string> ExtractAsync(string operation, string input) =>
        Task.Run(() => Extract(operation, input));

    private static string Extract(string operation, string input)
    {
        FileProcessor.EnsureSourceExists(input);
        var text = operation switch
        {
            "pdf_text" => PdfTextProcessor.ReadText(input),
            "docx_text" => OfficeDocumentProcessor.ReadText(input),
            "image_ocr" => OcrProcessor.ReadText(input),
            _ => throw new InvalidOperationException("Cette opération d’extraction n’est pas prise en charge."),
        };
        return text.Trim();
    }
}
