using System;
using System.Threading;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.Windows.AppLifecycle;
using WinRT;

namespace LumiCanvas;

public static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        EnsureWorkingDirectory();
        XamlCheckProcessRequirements();
        ComWrappersSupport.InitializeComWrappers();

        var mainInstance = AppInstance.FindOrRegisterForKey("main");
        var currentInstance = AppInstance.GetCurrent();

        if (!mainInstance.IsCurrent)
        {
            mainInstance.RedirectActivationToAsync(currentInstance.GetActivatedEventArgs()).AsTask().GetAwaiter().GetResult();
            return;
        }

        Application.Start(callbackParams =>
        {
            var dispatcherQueue = DispatcherQueue.GetForCurrentThread();
            SynchronizationContext.SetSynchronizationContext(new DispatcherQueueSynchronizationContext(dispatcherQueue));
            _ = new App();
        });
    }

    private static void EnsureWorkingDirectory()
    {
        try
        {
            var executablePath = Environment.ProcessPath;
            if (!string.IsNullOrWhiteSpace(executablePath))
            {
                var executableDirectory = System.IO.Path.GetDirectoryName(executablePath);
                if (!string.IsNullOrWhiteSpace(executableDirectory) && System.IO.Directory.Exists(executableDirectory))
                {
                    System.IO.Directory.SetCurrentDirectory(executableDirectory);
                }
            }
        }
        catch
        {
        }
    }

    [System.Runtime.InteropServices.DllImport("Microsoft.ui.xaml.dll")]
    private static extern void XamlCheckProcessRequirements();
}
