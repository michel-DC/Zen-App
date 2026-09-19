using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Zen.WinUI;

public sealed partial class MainWindow : Window
{
    private static readonly System.Collections.Generic.IReadOnlyDictionary<string, ToolDefinition> tools = ToolCatalog.All;

    private string operation = "docx_pdf";
    private string? inputPath;
    private string? outputPath;
    private bool isBusy;

    public MainWindow()
    {
        InitializeComponent();
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "Zen.ico"));
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1280, 800));
        SetOperation(operation);
        Navigation.SelectedItem = DocxPdfItem;
    }

    private void Navigation_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (isBusy) return;
        if (args.SelectedItem is NavigationViewItem item && item.Tag is string id && tools.ContainsKey(id)) SetOperation(id);
    }

    private void SetOperation(string id)
    {
        operation = id; inputPath = null; outputPath = null;
        var tool = tools[id];
        PageTitle.Text = tool.Title; PageDescription.Text = tool.Description;
        SourcePath.Text = string.Empty; OutputPath.Text = string.Empty;
        ExtractedTextBox.Text = string.Empty;
        CopyTextButton.IsEnabled = false;
        OutputHint.Text = "Zen suggérera un fichier de sortie dans le dossier source.";
        ResultLabel.Text = "Choisissez un fichier source pour commencer.";
        RunButton.Content = ToolCatalog.GetActionLabel(id);
        RunButton.IsEnabled = false;
        var extractsText = tool.ExtractsText;
        FilesDescription.Text = extractsText
            ? "Sélectionnez le fichier dont vous souhaitez lire et copier le texte."
            : "Sélectionnez la source puis confirmez l’emplacement du résultat.";
        OutputDivider.Visibility = extractsText ? Visibility.Collapsed : Visibility.Visible;
        OutputSection.Visibility = extractsText ? Visibility.Collapsed : Visibility.Visible;
        ExtractedTextSection.Visibility = extractsText ? Visibility.Visible : Visibility.Collapsed;
        var crop = id == "crop"; var rounded = id == "rounded";
        OptionsCard.Visibility = crop || rounded ? Visibility.Visible : Visibility.Collapsed;
        CropOptions.Visibility = crop ? Visibility.Visible : Visibility.Collapsed;
        RadiusOptions.Visibility = rounded ? Visibility.Visible : Visibility.Collapsed;
        StatusInfo.IsOpen = false;
    }

    private async void ChooseInput_Click(object sender, RoutedEventArgs e)
    {
        if (isBusy) return;
        var picker = new FileOpenPicker();
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
        foreach (var extension in tools[operation].Extensions) picker.FileTypeFilter.Add(extension);
        var file = await picker.PickSingleFileAsync();
        if (file is null) return;
        inputPath = file.Path; outputPath = null; SourcePath.Text = inputPath; OutputPath.Text = string.Empty;
        ClearExtractedText();
        if (!tools[operation].ExtractsText)
            OutputHint.Text = $"Nom suggéré : {Path.GetFileNameWithoutExtension(inputPath)}-zen.{tools[operation].Extension}";
        ResultLabel.Text = $"Prêt : {Path.GetFileName(inputPath)}";
        RunButton.IsEnabled = true;
        ShowStatus("Fichier sélectionné.", InfoBarSeverity.Success);
    }

    private void SourcePath_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (isBusy) return;
        ClearExtractedText();
        var candidate = SourcePath.Text.Trim().Trim('"');
        if (candidate.Length == 0)
        {
            inputPath = null;
            outputPath = null;
            RunButton.IsEnabled = false;
            return;
        }

        var extension = Path.GetExtension(candidate);
        var isSupported = tools[operation].Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(candidate) || !isSupported)
        {
            inputPath = null;
            outputPath = null;
            RunButton.IsEnabled = false;
            ResultLabel.Text = isSupported ? "Ce fichier est introuvable." : "Ce format ne correspond pas à l’outil sélectionné.";
            return;
        }

        inputPath = candidate;
        outputPath = null;
        OutputPath.Text = string.Empty;
        ClearExtractedText();
        if (!tools[operation].ExtractsText)
            OutputHint.Text = $"Nom suggéré : {Path.GetFileNameWithoutExtension(inputPath)}-zen.{tools[operation].Extension}";
        ResultLabel.Text = $"Prêt : {Path.GetFileName(inputPath)}";
        RunButton.IsEnabled = true;
    }

    private async void ChooseOutput_Click(object sender, RoutedEventArgs e)
    {
        if (isBusy) return;
        if (tools[operation].ExtractsText) return;
        if (inputPath is null) { ShowStatus("Choisissez d’abord un fichier source.", InfoBarSeverity.Warning); return; }
        var picker = new FileSavePicker { SuggestedFileName = $"{Path.GetFileNameWithoutExtension(inputPath)}-zen", SuggestedStartLocation = PickerLocationId.DocumentsLibrary };
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
        picker.FileTypeChoices.Add(tools[operation].Extension.ToUpperInvariant(), new[] { $".{tools[operation].Extension}" });
        var file = await picker.PickSaveFileAsync();
        if (file is null) return;
        outputPath = file.Path; OutputPath.Text = outputPath; OutputHint.Text = "Le résultat sera enregistré à cet emplacement.";
        ShowStatus("Emplacement de sortie sélectionné.", InfoBarSeverity.Informational);
    }

    private async void Run_Click(object sender, RoutedEventArgs e)
    {
        if (inputPath is null) { ShowStatus("Choisissez un fichier source pour continuer.", InfoBarSeverity.Warning); return; }

        var jobOperation = operation;
        var jobInput = inputPath;
        var jobOutput = outputPath;
        var tool = tools[jobOperation];
        var cropX = (int)CropX.Value;
        var cropY = (int)CropY.Value;
        var cropWidth = (int)CropWidth.Value;
        var cropHeight = (int)CropHeight.Value;
        var radius = (int)Radius.Value;

        SetBusy(true);
        try
        {
            if (tool.ExtractsText)
            {
                var text = await TextExtractionService.ExtractAsync(jobOperation, jobInput);
                ExtractedTextBox.Text = text;
                CopyTextButton.IsEnabled = true;
                ResultLabel.Text = $"Texte extrait : {text.Length:N0} caractères.";
                ShowStatus("Le texte a été extrait localement.", InfoBarSeverity.Success);
            }
            else
            {
                var result = await FileProcessor.ProcessAsync(jobOperation, jobInput, jobOutput, tool.Extension, cropX, cropY, cropWidth, cropHeight, radius);
                outputPath = result; OutputPath.Text = result; OutputHint.Text = "Traitement terminé."; ResultLabel.Text = $"Créé : {Path.GetFileName(result)}";
                ShowStatus("Le fichier a été créé localement.", InfoBarSeverity.Success);
            }
        }
        catch (Exception exception) { ShowStatus(exception.Message, InfoBarSeverity.Error); }
        finally { SetBusy(false); }
    }

    private void SetBusy(bool busy)
    {
        isBusy = busy;
        RunButton.IsEnabled = !busy && inputPath is not null;
        BrowseInputButton.IsEnabled = !busy;
        BrowseOutputButton.IsEnabled = !busy && !tools[operation].ExtractsText;
        CropX.IsEnabled = !busy;
        CropY.IsEnabled = !busy;
        CropWidth.IsEnabled = !busy;
        CropHeight.IsEnabled = !busy;
        Radius.IsEnabled = !busy;
        foreach (var menuItem in Navigation.MenuItems)
            if (menuItem is NavigationViewItem item) item.IsEnabled = !busy;
        Progress.IsActive = busy;
        Progress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
    }
    private void CopyText_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(ExtractedTextBox.Text)) return;
        var package = new DataPackage();
        package.SetText(ExtractedTextBox.Text);
        Clipboard.SetContent(package);
        Clipboard.Flush();
        ShowStatus("Texte copié dans le presse-papiers.", InfoBarSeverity.Success);
    }

    private void ClearExtractedText()
    {
        ExtractedTextBox.Text = string.Empty;
        CopyTextButton.IsEnabled = false;
    }

    private void ShowStatus(string message, InfoBarSeverity severity) { StatusInfo.Message = message; StatusInfo.Severity = severity; StatusInfo.IsOpen = true; }
}
