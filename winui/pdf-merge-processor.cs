using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Zen.WinUI;

internal sealed record PdfMergeResult(string OutputPath, int FileCount, int PageCount);

internal static class PdfMergeProcessor
{
    internal static Task<PdfMergeResult> MergeAsync(IReadOnlyList<string> inputs, string? requestedOutput)
    {
        var files = inputs.ToArray();
        return Task.Run(() => Merge(files, requestedOutput));
    }

    internal static string SuggestOutputPath(string input) => OutputFileNaming.SuggestPath(input, "pdf_merge", "pdf");

    private static PdfMergeResult Merge(string[] inputs, string? requestedOutput)
    {
        if (inputs.Length < 2)
            throw new InvalidOperationException("Sélectionnez au moins deux fichiers PDF à fusionner.");

        var files = inputs.Select(Path.GetFullPath).ToArray();
        var output = Path.GetFullPath(requestedOutput ?? SuggestOutputPath(files[0]));
        if (!string.Equals(Path.GetExtension(output), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Choisissez un fichier de sortie au format PDF.");
        if (files.Contains(output, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("Le fichier de sortie doit être différent de tous les PDF sources.");

        using var merged = new PdfDocument();
        foreach (var file in files)
        {
            if (!string.Equals(Path.GetExtension(file), ".pdf", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"« {Path.GetFileName(file)} » n’est pas un fichier PDF.");
            if (!File.Exists(file))
                throw new FileNotFoundException($"Le PDF « {Path.GetFileName(file)} » est introuvable. Sélectionnez-le à nouveau.", file);
            try
            {
                using var document = PdfReader.Open(file, PdfDocumentOpenMode.Import);
                if (document.PageCount == 0)
                    throw new InvalidOperationException("Ce PDF ne contient aucune page.");
                foreach (var page in document.Pages)
                    PdfA4PageNormalizer.Normalize(merged.AddPage(page));
            }
            catch (Exception exception) when (exception is not OutOfMemoryException)
            {
                throw new InvalidOperationException(
                    $"Impossible de lire « {Path.GetFileName(file)} ». Vérifiez que le PDF est valide et accessible sans mot de passe.", exception);
            }
        }

        var outputFolder = Path.GetDirectoryName(output)!;
        Directory.CreateDirectory(outputFolder);
        var temporaryOutput = Path.Combine(outputFolder, $".zen-merge-{Guid.NewGuid():N}.pdf");
        var pageCount = merged.PageCount;
        try
        {
            merged.Save(temporaryOutput);

            using (var document = PdfReader.Open(temporaryOutput, PdfDocumentOpenMode.Import))
                if (document.PageCount != pageCount)
                    throw new InvalidOperationException("La fusion n’a pas conservé toutes les pages. Les fichiers sources sont inchangés.");

            File.Move(temporaryOutput, output, true);
            return new PdfMergeResult(output, files.Length, pageCount);
        }
        finally
        {
            if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
        }
    }
}
