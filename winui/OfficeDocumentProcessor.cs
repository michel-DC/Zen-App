using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace Zen.WinUI;

internal static class OfficeDocumentProcessor
{
    internal static string ReadText(string input)
    {
        using var archive = ZipFile.OpenRead(input);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidOperationException("Le document DOCX ne contient pas de texte lisible.");
        using var stream = entry.Open();
        var document = XDocument.Load(stream);
        var paragraphs = document.Descendants()
            .Where(node => node.Name.LocalName == "p")
            .Select(node => node.Value);
        var text = string.Join(Environment.NewLine, paragraphs).Trim();
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("Le document DOCX ne contient pas de texte lisible.");
        return text;
    }

    internal static void ConvertToPdf(string input, string output)
    {
        if (WordToPdfConverter.IsAvailable)
        {
            WordToPdfConverter.Convert(input, output);
            return;
        }

        ConvertWithLibreOffice(input, output);
    }

    private static void ConvertWithLibreOffice(string input, string output)
    {
        var conversionFolder = Path.Combine(Path.GetTempPath(), $"zen-libreoffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(conversionFolder);
        try
        {
            using var process = StartLibreOffice(input, conversionFolder);
            process.WaitForExit();
            CopyConvertedPdf(process, input, output, conversionFolder);
        }
        finally
        {
            if (Directory.Exists(conversionFolder)) Directory.Delete(conversionFolder, true);
        }
    }

    private static Process StartLibreOffice(string input, string outputFolder)
    {
        var startInfo = new ProcessStartInfo(FindLibreOffice())
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("--headless");
        startInfo.ArgumentList.Add("--convert-to");
        startInfo.ArgumentList.Add("pdf");
        startInfo.ArgumentList.Add("--outdir");
        startInfo.ArgumentList.Add(outputFolder);
        startInfo.ArgumentList.Add(input);
        return Process.Start(startInfo) ?? throw new InvalidOperationException("Impossible de démarrer LibreOffice.");
    }

    private static void CopyConvertedPdf(Process process, string input, string output, string folder)
    {
        var generated = Path.Combine(folder, $"{Path.GetFileNameWithoutExtension(input)}.pdf");
        if (process.ExitCode != 0 || !File.Exists(generated))
            throw new InvalidOperationException("LibreOffice n’a pas pu convertir ce document Word en PDF.");
        File.Copy(generated, output, true);
    }

    private static string FindLibreOffice()
    {
        var candidates = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.com",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.com"
        };
        return candidates.FirstOrDefault(File.Exists)
            ?? throw new InvalidOperationException("LibreOffice est introuvable. Installez-le puis relancez Zen.");
    }
}
