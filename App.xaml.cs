using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Windows.AppLifecycle;
using Microsoft.UI.Xaml;

namespace LumiCanvas
{
    public partial class App : Application
    {
        private MainWindow? _window;
        private readonly WorkspaceSession _session = new();

        public App()
        {
            InitializeComponent();
            AppInstance.GetCurrent().Activated += AppInstance_Activated;
            UnhandledException += App_UnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
        {
            WriteDiagnostic("App.UnhandledException", e.Exception);
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, System.UnhandledExceptionEventArgs e)
        {
            WriteDiagnostic("AppDomain.UnhandledException", e.ExceptionObject as Exception);
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            WriteDiagnostic("TaskScheduler.UnobservedTaskException", e.Exception);
            e.SetObserved();
        }

        internal static void WriteDiagnostic(string source, Exception? exception)
        {
            try
            {
                var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LumiCanvas", "Logs");
                Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "runtime.log");

                var builder = new StringBuilder();
                builder.AppendLine("========================================");
                builder.AppendLine(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss.fff zzz"));
                builder.AppendLine(source);
                if (exception is not null)
                {
                    builder.AppendLine(exception.ToString());
                }

                File.AppendAllText(logPath, builder.ToString(), Encoding.UTF8);
            }
            catch
            {
            }
        }

        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            EnsureWindow();
            HandleActivation(AppInstance.GetCurrent().GetActivatedEventArgs());
        }

        private void EnsureWindow()
        {
            if (_window is not null)
            {
                return;
            }

            _window = new MainWindow(_session);
            _window.Activate();
            _window.HideToBackground();
        }

        private void AppInstance_Activated(object? sender, AppActivationArguments args)
        {
            EnsureWindow();
            _window?.DispatcherQueue.TryEnqueue(() => HandleActivation(args));
        }

        private void HandleActivation(AppActivationArguments args)
        {
            if (args.Kind == ExtendedActivationKind.Protocol && args.Data is Windows.ApplicationModel.Activation.ProtocolActivatedEventArgs protocolArgs)
            {
                HandleProtocolActivation(protocolArgs.Uri);
            }
        }

        private void HandleProtocolActivation(Uri? uri)
        {
            if (_window is null)
            {
                return;
            }

            if (TryGetTaskIdFromUri(uri, out var taskId) && _window.TryOpenTask(taskId))
            {
                return;
            }

            _window.ShowSidebarFromProtocol("未找到协议链接对应的任务。", uri is not null);
        }

        private static bool TryGetTaskIdFromUri(Uri? uri, out Guid taskId)
        {
            taskId = Guid.Empty;
            if (uri is null)
            {
                return false;
            }

            if (string.Equals(uri.Host, "task", StringComparison.OrdinalIgnoreCase))
            {
                return Guid.TryParse(uri.AbsolutePath.Trim('/'), out taskId);
            }

            var query = uri.Query.TrimStart('?');
            if (string.IsNullOrWhiteSpace(query))
            {
                return false;
            }

            foreach (var segment in query.Split('&', StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = segment.Split('=', 2, StringSplitOptions.TrimEntries);
                if (parts.Length == 2 && string.Equals(parts[0], "taskId", StringComparison.OrdinalIgnoreCase))
                {
                    return Guid.TryParse(Uri.UnescapeDataString(parts[1]), out taskId);
                }
            }

            return false;
        }
    }
}
