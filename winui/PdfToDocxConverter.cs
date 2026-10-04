using System;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace Zen.WinUI;

internal static class PdfToDocxConverter
{
    internal static void Convert(string input, string output)
    {
        Exception? conversionError = null;
        var conversionThread = new Thread(() => RunWordConversion(input, output, out conversionError))
        {
            Name = "Zen PDF Reflow"
        };
        conversionThread.SetApartmentState(ApartmentState.STA);
        conversionThread.Start();
        conversionThread.Join();
        if (conversionError is not null) ExceptionDispatchInfo.Capture(conversionError).Throw();
    }

    private static void RunWordConversion(string input, string output, out Exception? conversionError)
    {
        try
        {
            ConvertWithWord(input, output);
            conversionError = null;
        }
        catch (Exception exception)
        {
            conversionError = exception;
        }
    }

    private static void ConvertWithWord(string input, string output)
    {
        object? applicationObject = null;
        object? documentsObject = null;
        object? documentObject = null;
        var temporaryOutput = CreateTemporaryOutputPath(output);
        var temporaryInput = Path.Combine(Path.GetTempPath(), $"zen-pdf-{Guid.NewGuid():N}.pdf");

        try
        {
            File.Copy(input, temporaryInput, true);
            using var fonts = TemporaryFontLoader.Load(
                Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts"));
            applicationObject = CreateWordApplication();
            dynamic application = applicationObject;
            var caption = ConfigureWord(application);
            documentsObject = application.Documents;
            documentObject = OpenPdf(application, documentsObject, temporaryInput, caption);
            SaveAsDocx(documentObject, temporaryOutput);
            ReleaseComObject(ref documentObject);
            var sourcePageCount = DocxLayoutRestorer.Restore(input, temporaryOutput);
            if (sourcePageCount == 1) FitSinglePage(documentsObject, temporaryOutput);
            CloseWord(ref documentObject, ref documentsObject, ref applicationObject);
            MoveValidatedOutput(temporaryOutput, output);
        }
        catch (COMException exception)
        {
            throw new InvalidOperationException($"Microsoft Word n’a pas pu reconstruire ce PDF : {exception.Message}", exception);
        }
        finally
        {
            CleanupWord(ref documentObject, ref documentsObject, ref applicationObject);
            DeleteIfExists(temporaryOutput);
            DeleteIfExists(temporaryInput);
        }
    }

    private static object CreateWordApplication()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word est requis pour conserver la mise en page lors de la conversion PDF en DOCX.");
        return Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Impossible de démarrer Microsoft Word.");
    }

    private static string ConfigureWord(dynamic application)
    {
        var caption = $"Zen PDF {Guid.NewGuid():N}";
        application.Caption = caption;
        application.Visible = true;
        application.DisplayAlerts = 0;
        application.AutomationSecurity = 3;
        application.Options.DoNotPromptForConvert = true;
        application.Options.ConfirmConversions = false;
        return caption;
    }

    private static object OpenPdf(dynamic application, object documentsObject, string input, string caption)
    {
        dynamic documents = documentsObject;
        using var warningHelper = WordImportWarningHelper.Start(caption);
        try
        {
            var document = documents.Open(
                FileName: input,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Revert: false,
                Visible: false,
                OpenAndRepair: false,
                NoEncodingDialog: true);
            application.Visible = false;
            return document;
        }
        finally
        {
            if (!warningHelper.WaitForExit(2000)) warningHelper.Kill(true);
        }
    }

    private static void SaveAsDocx(object documentObject, string output)
    {
        dynamic document = documentObject;
        document.EmbedTrueTypeFonts = true;
        document.SaveSubsetFonts = true;
        document.SaveAs2(
            FileName: output,
            FileFormat: 16,
            AddToRecentFiles: false,
            EmbedTrueTypeFonts: true,
            CompatibilityMode: 65535);
        document.Close(false);
    }

    private static void FitSinglePage(object documentsObject, string path)
    {
        dynamic documents = documentsObject;
        object? documentObject = null;
        try
        {
            documentObject = documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            dynamic document = documentObject;
            document.Repaginate();
            var pages = (int)document.ComputeStatistics(2);
            for (var attempt = 0; pages > 1 && attempt < 3; attempt++)
            {
                document.FitToPages();
                document.Repaginate();
                var reducedPages = (int)document.ComputeStatistics(2);
                if (reducedPages >= pages) break;
                pages = reducedPages;
            }
            if (pages > 1)
                throw new InvalidOperationException(
                    "Word n’a pas pu conserver le contenu de ce PDF d’une page sur une seule page DOCX.");
            document.Save();
            document.Close(false);
        }
        finally
        {
            if (documentObject is not null)
            {
                try { ((dynamic)documentObject).Close(false); }
                catch { }
            }
            ReleaseComObject(ref documentObject);
        }
    }

    private static void MoveValidatedOutput(string temporaryOutput, string output)
    {
        if (!File.Exists(temporaryOutput) || new FileInfo(temporaryOutput).Length == 0)
            throw new InvalidOperationException("Microsoft Word n’a produit aucun document.");
        File.Move(temporaryOutput, output, true);
    }

    private static string CreateTemporaryOutputPath(string output) =>
        Path.Combine(
            Path.GetDirectoryName(output)!,
            $".{Path.GetFileNameWithoutExtension(output)}-{Guid.NewGuid():N}.docx");

    private static void CloseWord(ref object? document, ref object? documents, ref object? application)
    {
        ReleaseComObject(ref document);
        ((dynamic)application!).Quit(false);
        ReleaseComObject(ref documents);
        ReleaseComObject(ref application);
    }

    private static void CleanupWord(ref object? document, ref object? documents, ref object? application)
    {
        if (document is not null)
        {
            try { ((dynamic)document).Close(false); }
            catch { }
        }
        ReleaseComObject(ref document);
        if (application is not null)
        {
            try { ((dynamic)application).Quit(false); }
            catch { }
        }
        ReleaseComObject(ref documents);
        ReleaseComObject(ref application);
    }

    private static void ReleaseComObject(ref object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        value = null;
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
