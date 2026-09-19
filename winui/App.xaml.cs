using System;
using System.IO;
using Microsoft.UI.Xaml;

namespace Zen.WinUI;

public partial class App : Application
{
    private Window? window;

    public App()
    {
        UnhandledException += (_, args) => WriteStartupFailure(args.Exception);

        try
        {
            InitializeComponent();
        }
        catch (Exception exception)
        {
            WriteStartupFailure(exception);
            throw;
        }
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        try
        {
            window = new MainWindow();
            window.Activate();
        }
        catch (Exception exception)
        {
            WriteStartupFailure(exception);
            throw;
        }
    }

    private static void WriteStartupFailure(Exception exception)
    {
        File.WriteAllText(Path.Combine(Path.GetTempPath(), "zen-winui-startup-error.txt"), exception.ToString());
    }
}
