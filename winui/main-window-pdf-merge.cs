using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Zen.WinUI;

internal sealed record PdfMergeInput(string Path)
{
    public string FileName => System.IO.Path.GetFileName(Path);
}

public sealed partial class MainWindow
{
    private readonly ObservableCollection<PdfMergeInput> mergeInputs = new();
    private bool IsPdfMerge => operation == "pdf_merge";
    private bool HasValidInput => IsPdfMerge ? mergeInputs.Count >= 2 : inputPath is not null;

    private async void ChooseMergeInputs_Click(object sender, RoutedEventArgs e)
    {
        if (isBusy) return;
        var picker = new FileOpenPicker();
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
        picker.FileTypeFilter.Add(".pdf");
        try
        {
            var files = await picker.PickMultipleFilesAsync();
            if (files.Count == 0) return;
            foreach (var file in files)
                if (!mergeInputs.Any(input => string.Equals(input.Path, file.Path, StringComparison.OrdinalIgnoreCase)))
                    mergeInputs.Add(new PdfMergeInput(file.Path));
            RefreshMergeInputs();
            ShowStatus($"{mergeInputs.Count} PDF dans la liste. Vérifiez leur ordre avant de fusionner.", InfoBarSeverity.Informational);
        }
        catch (Exception exception) { ShowStatus(exception.Message, InfoBarSeverity.Error); }
    }

    private void MergeFiles_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateMergeButtons();

    private void MoveMergeUp_Click(object sender, RoutedEventArgs e) => MoveMergeInput(-1);

    private void MoveMergeDown_Click(object sender, RoutedEventArgs e) => MoveMergeInput(1);

    private void MoveMergeInput(int direction)
    {
        var index = MergeFilesList.SelectedIndex;
        var target = index + direction;
        if (isBusy || index < 0 || target < 0 || target >= mergeInputs.Count) return;
        mergeInputs.Move(index, target);
        MergeFilesList.SelectedIndex = target;
        RefreshMergeInputs();
    }

    private void RemoveMergeInput_Click(object sender, RoutedEventArgs e)
    {
        var index = MergeFilesList.SelectedIndex;
        if (isBusy || index < 0) return;
        mergeInputs.RemoveAt(index);
        MergeFilesList.SelectedIndex = Math.Min(index, mergeInputs.Count - 1);
        RefreshMergeInputs();
    }

    private void RefreshMergeInputs()
    {
        inputPath = mergeInputs.FirstOrDefault()?.Path;
        MergeEmptyHint.Visibility = mergeInputs.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        ResultLabel.Text = mergeInputs.Count < 2
            ? "Sélectionnez au moins deux PDF pour commencer."
            : $"{mergeInputs.Count} PDF seront réunis dans l’ordre affiché.";
        if (outputPath is null)
            OutputHint.Text = inputPath is null
                ? "Le PDF fusionné sera proposé dans le dossier du premier fichier."
                : $"Nom suggéré : {Path.GetFileName(PdfMergeProcessor.SuggestOutputPath(inputPath))}";
        RunButton.IsEnabled = !isBusy && HasValidInput;
        UpdateMergeButtons();
    }

    private void UpdateMergeButtons()
    {
        var index = MergeFilesList.SelectedIndex;
        AddMergeInputsButton.IsEnabled = !isBusy;
        MergeFilesList.IsEnabled = !isBusy;
        MoveMergeUpButton.IsEnabled = !isBusy && index > 0;
        MoveMergeDownButton.IsEnabled = !isBusy && index >= 0 && index < mergeInputs.Count - 1;
        RemoveMergeInputButton.IsEnabled = !isBusy && index >= 0;
    }

    private async Task RunPdfMergeAsync()
    {
        if (!HasValidInput)
        {
            ShowStatus("Sélectionnez au moins deux PDF pour continuer.", InfoBarSeverity.Warning);
            return;
        }
        var files = mergeInputs.Select(input => input.Path).ToArray();
        var requestedOutput = outputPath;
        SetBusy(true);
        ResultLabel.Text = $"Fusion de {files.Length} PDF en cours…";
        try
        {
            var result = await PdfMergeProcessor.MergeAsync(files, requestedOutput);
            outputPath = result.OutputPath;
            OutputPath.Text = result.OutputPath;
            OutputHint.Text = "Fusion terminée. Les fichiers sources sont conservés.";
            ResultLabel.Text = $"{result.FileCount} PDF fusionnés · {result.PageCount} pages.";
            ShowStatus($"PDF créé : {Path.GetFileName(result.OutputPath)}", InfoBarSeverity.Success);
        }
        catch (Exception exception)
        {
            ResultLabel.Text = "La fusion n’a pas abouti. Vérifiez les fichiers puis réessayez.";
            ShowStatus(exception.Message, InfoBarSeverity.Error);
        }
        finally { SetBusy(false); }
    }
}
