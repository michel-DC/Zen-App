using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace Zen.WinUI;

internal sealed class TemporaryFontLoader : IDisposable
{
    private readonly List<string> loadedFonts = new();

    private TemporaryFontLoader()
    {
    }

    internal static TemporaryFontLoader Load(string directory)
    {
        var loader = new TemporaryFontLoader();
        if (!Directory.Exists(directory)) return loader;

        foreach (var path in Directory.EnumerateFiles(directory, "*.ttf"))
        {
            if (AddFontResourceEx(path, 0, IntPtr.Zero) > 0)
                loader.loadedFonts.Add(path);
        }

        return loader;
    }

    public void Dispose()
    {
        foreach (var path in loadedFonts)
            RemoveFontResourceEx(path, 0, IntPtr.Zero);
        loadedFonts.Clear();
    }

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int AddFontResourceEx(string name, uint flags, IntPtr reserved);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool RemoveFontResourceEx(string name, uint flags, IntPtr reserved);
}
