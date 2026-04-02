using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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

            if (_window is not null && TryExtractLumiPath(args.Arguments, out var launchArchivePath))
            {
                if (_window.TryOpenTaskArchive(launchArchivePath))
                {
                    return;
                }

                _window.ShowSidebarFromProtocol("未能打开指定的 .lumi 存档。", true);
                WriteDiagnostic("App.OnLaunched.OpenArchiveFailed", new InvalidOperationException($"Path={launchArchivePath}"));
                return;
            }

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
                return;
            }

            if (TryGetArchivePathFromActivation(args, out var archivePath) && _window is not null)
            {
                if (_window.TryOpenTaskArchive(archivePath))
                {
                    return;
                }

                _window.ShowSidebarFromProtocol("未能打开指定的 .lumi 存档。", true);
                WriteDiagnostic("App.HandleActivation.OpenArchiveFailed", new InvalidOperationException($"Path={archivePath}"));
                return;
            }

            WriteDiagnostic("App.HandleActivation.NoArchivePath", new InvalidOperationException($"Kind={args.Kind}; DataType={args.Data?.GetType().FullName ?? "null"}"));
        }

        private static bool TryGetArchivePathFromActivation(AppActivationArguments args, out string archivePath)
        {
            archivePath = string.Empty;

            if (args.Kind == ExtendedActivationKind.File && args.Data is Windows.ApplicationModel.Activation.IFileActivatedEventArgs fileArgs)
            {
                var file = fileArgs.Files.OfType<Windows.Storage.StorageFile>().FirstOrDefault();
                if (file is not null && string.Equals(Path.GetExtension(file.Path), ".lumi", StringComparison.OrdinalIgnoreCase))
                {
                    archivePath = file.Path;
                    return true;
                }
            }

            if (args.Kind == ExtendedActivationKind.Launch && args.Data is Windows.ApplicationModel.Activation.ILaunchActivatedEventArgs launchArgs)
            {
                var raw = launchArgs.Arguments?.Trim();
                if (TryExtractLumiPath(raw, out var launchPath))
                {
                    archivePath = launchPath;
                    return true;
                }
            }

            var commandLineArgs = Environment.GetCommandLineArgs();
            foreach (var arg in commandLineArgs.Skip(1))
            {
                if (TryExtractLumiPath(arg, out var cliPath))
                {
                    archivePath = cliPath;
                    return true;
                }
            }

            return false;
        }

        private static bool TryExtractLumiPath(string? raw, out string archivePath)
        {
            archivePath = string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            static bool IsLumiPath(string candidate)
            {
                return !string.IsNullOrWhiteSpace(candidate) &&
                       string.Equals(Path.GetExtension(candidate), ".lumi", StringComparison.OrdinalIgnoreCase);
            }

            static bool TryNormalize(string candidate, out string normalized)
            {
                normalized = string.Empty;
                if (!IsLumiPath(candidate))
                {
                    return false;
                }

                if (candidate.Contains('"'))
                {
                    return false;
                }

                try
                {
                    normalized = Path.GetFullPath(candidate.Trim());
                    return File.Exists(normalized);
                }
                catch
                {
                    return false;
                }
            }

            var trimmed = raw.Trim().Trim('"');
            if (TryNormalize(trimmed, out var normalizedTrimmed))
            {
                archivePath = normalizedTrimmed;
                return true;
            }

            var text = raw.Trim();
            var regexMatches = Regex.Matches(text, "[A-Za-z]:\\\\[^\"\r\n]*?\\.lumi", RegexOptions.IgnoreCase);
            for (var i = regexMatches.Count - 1; i >= 0; i--)
            {
                if (TryNormalize(regexMatches[i].Value, out var regexPath))
                {
                    archivePath = regexPath;
                    return true;
                }
            }

            var index = 0;
            while (index < text.Length)
            {
                if (text[index] == '"')
                {
                    var end = text.IndexOf('"', index + 1);
                    if (end > index)
                    {
                        var quoted = text[(index + 1)..end].Trim();
                        if (TryNormalize(quoted, out var quotedPath))
                        {
                            archivePath = quotedPath;
                            return true;
                        }

                        index = end + 1;
                        continue;
                    }
                }

                index++;
            }

            foreach (var token in text.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var cleaned = token.Trim('"');
                if (TryNormalize(cleaned, out var tokenPath))
                {
                    archivePath = tokenPath;
                    return true;
                }
            }

            return false;
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
