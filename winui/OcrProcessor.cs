using System;
using System.IO;
using Tesseract;

namespace Zen.WinUI;

internal static class OcrProcessor
{
    internal static string ReadText(string input)
    {
        var dataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        EnsureLanguageDataExists(dataPath);
        try
        {
            using var engine = new TesseractEngine(dataPath, "fra+eng", EngineMode.LstmOnly);
            using var image = Pix.LoadFromFile(input);
            using var page = engine.Process(image);
            var text = page.GetText().Trim();
            if (string.IsNullOrWhiteSpace(text))
                throw new InvalidOperationException("Aucun texte n’a été détecté dans cette image.");
            return text;
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException($"La lecture OCR a échoué : {exception.Message}", exception);
        }
    }

    private static void EnsureLanguageDataExists(string dataPath)
    {
        var hasFrench = File.Exists(Path.Combine(dataPath, "fra.traineddata"));
        var hasEnglish = File.Exists(Path.Combine(dataPath, "eng.traineddata"));
        if (!hasFrench || !hasEnglish)
            throw new InvalidOperationException("Les données OCR intégrées sont introuvables. Réinstallez Zen.");
    }
}
