using System;
using System.IO;
using System.Threading.Tasks;

namespace Zen.WinUI;

internal static class FileProcessor
{
    internal static Task<string> ProcessAsync(
        string operation,
        string input,
        string? requestedOutput,
        string extension,
        int x,
        int y,
        int width,
        int height,
        int radius) =>
        Task.Run(() => Process(operation, input, requestedOutput, extension, x, y, width, height, radius));

    private static string Process(
        string operation,
        string input,
        string? requestedOutput,
        string extension,
        int x,
        int y,
        int width,
        int height,
        int radius)
    {
        EnsureSourceExists(input);
        var output = ResolveOutputPath(input, requestedOutput, extension);
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        switch (operation)
        {
            case "docx_pdf": OfficeDocumentProcessor.ConvertToPdf(input, output); break;
            case "pdf_docx": PdfToDocxConverter.Convert(input, output); break;
            default: ImageProcessor.Process(operation, input, output, new CropArea(x, y, width, height), radius); break;
        }
        return output;
    }

    internal static void EnsureSourceExists(string input)
    {
        if (!File.Exists(input)) throw new FileNotFoundException("Le fichier source est introuvable.", input);
    }

    private static string ResolveOutputPath(string input, string? requestedOutput, string extension)
    {
        var output = requestedOutput ?? Path.Combine(
            Path.GetDirectoryName(input)!,
            $"{Path.GetFileNameWithoutExtension(input)}-zen.{extension}");
        if (string.Equals(Path.GetFullPath(input), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Le fichier de sortie doit être différent du fichier source.");
        return output;
    }
}
