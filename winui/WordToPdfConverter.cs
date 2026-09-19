using System;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace Zen.WinUI;

internal static class WordToPdfConverter
{
    internal static bool IsAvailable => Type.GetTypeFromProgID("Word.Application") is not null;

    internal static void Convert(string input, string output)
    {
        Exception? conversionError = null;
        var thread = new Thread(() => RunConversion(input, output, out conversionError))
        {
            Name = "Zen Word PDF export"
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (conversionError is not null) ExceptionDispatchInfo.Capture(conversionError).Throw();
    }

    private static void RunConversion(string input, string output, out Exception? conversionError)
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
        var temporaryOutput = Path.Combine(
            Path.GetDirectoryName(output)!,
            $".{Path.GetFileNameWithoutExtension(output)}-{Guid.NewGuid():N}.pdf");

        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Microsoft Word est indisponible.");
            applicationObject = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Impossible de démarrer Microsoft Word.");
            dynamic application = applicationObject;
            application.Visible = false;
            application.DisplayAlerts = 0;
            application.AutomationSecurity = 3;
            documentsObject = application.Documents;
            dynamic documents = documentsObject;
            documentObject = documents.Open(
                FileName: input,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false,
                NoEncodingDialog: true);
            dynamic document = documentObject;
            document.Repaginate();
            document.ExportAsFixedFormat(
                OutputFileName: temporaryOutput,
                ExportFormat: 17,
                OpenAfterExport: false,
                OptimizeFor: 0,
                Range: 0,
                Item: 0,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: 1,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);
            document.Close(false);
            ReleaseComObject(ref documentObject);
            application.Quit(false);
            ReleaseComObject(ref documentsObject);
            ReleaseComObject(ref applicationObject);

            if (!File.Exists(temporaryOutput) || new FileInfo(temporaryOutput).Length == 0)
                throw new InvalidOperationException("Microsoft Word n’a produit aucun PDF.");
            File.Move(temporaryOutput, output, true);
        }
        catch (COMException exception)
        {
            throw new InvalidOperationException(
                $"Microsoft Word n’a pas pu créer le PDF : {exception.Message}",
                exception);
        }
        finally
        {
            CleanupWord(ref documentObject, ref documentsObject, ref applicationObject);
            if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
        }
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
}
