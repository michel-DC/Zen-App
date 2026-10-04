using System;
using System.IO;

namespace Zen.WinUI;

internal static class OutputFileNaming
{
    internal static string SuggestPath(string input, string operation, string extension)
    {
        var source = Path.GetFullPath(input);
        var folder = Path.GetDirectoryName(source)!;
        var name = Path.GetFileNameWithoutExtension(source);
        name = name.Replace("zen", string.Empty, StringComparison.OrdinalIgnoreCase).Trim(' ', '-', '_', '.');
        if (string.IsNullOrWhiteSpace(name)) name = "document";
        var suffix = operation switch
        {
            "pdf_merge" => "-fusion",
            "crop" => "-rognee",
            "rounded" => "-coins-arrondis",
            "circle" => "-circulaire",
            _ => string.Empty
        };
        var stem = name + suffix;
        var candidate = Path.Combine(folder, $"{stem}.{extension}");
        var index = 2;
        while (File.Exists(candidate) || Directory.Exists(candidate)
            || string.Equals(candidate, source, StringComparison.OrdinalIgnoreCase))
            candidate = Path.Combine(folder, $"{stem} ({index++}).{extension}");
        return candidate;
    }
}
