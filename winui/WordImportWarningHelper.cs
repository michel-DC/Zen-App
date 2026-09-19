using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Zen.WinUI;

internal static class WordImportWarningHelper
{
    internal static Process Start(string caption)
    {
        var encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(CreateScript(caption)));
        var powershell = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            @"WindowsPowerShell\v1.0\powershell.exe");
        if (!File.Exists(powershell))
            throw new InvalidOperationException("Le composant Windows PowerShell nécessaire à l’automatisation de Word est introuvable.");

        var startInfo = new ProcessStartInfo(powershell)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
        foreach (var argument in new[] { "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", encodedCommand })
            startInfo.ArgumentList.Add(argument);
        return Process.Start(startInfo)
            ?? throw new InvalidOperationException("Zen n’a pas pu démarrer l’assistant de conversion Word.");
    }

    private static string CreateScript(string caption)
    {
        var safeCaption = caption.Replace("'", "''", StringComparison.Ordinal);
        return $$"""
            Add-Type -AssemblyName UIAutomationClient
            Add-Type -AssemblyName UIAutomationTypes
            $deadline = (Get-Date).AddSeconds(30)
            do {
                $word = Get-Process WINWORD -ErrorAction SilentlyContinue |
                    Where-Object { $_.MainWindowTitle -like '*{{safeCaption}}*' } |
                    Select-Object -First 1
                if ($word -and $word.MainWindowHandle -ne 0) {
                    $root = [System.Windows.Automation.AutomationElement]::FromHandle($word.MainWindowHandle)
                    $condition = New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::NameProperty, 'OK')
                    $button = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
                    if ($button) {
                        $pattern = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                        $pattern.Invoke()
                        exit 0
                    }
                }
                Start-Sleep -Milliseconds 100
            } while ((Get-Date) -lt $deadline)
            exit 1
            """;
    }
}
