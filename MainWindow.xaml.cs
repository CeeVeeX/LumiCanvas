using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.Win32;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.System;
using WinRT.Interop;

namespace LumiCanvas;

public sealed partial class MainWindow : Window
{
    private const int HotKeyId = 0x4C43;
    private const string StartupRegistryValueName = "LumiCanvas";
    private const uint ModControl = 0x0002;
    private const uint ModNoRepeat = 0x4000;
    private const uint WmCommand = 0x0111;
    private const uint WmHotKey = 0x0312;
    private const uint WmTrayIcon = 0x8001;
    private const uint WmLButtonDblClk = 0x0203;
    private const uint WmRButtonUp = 0x0205;
    private const int GwlStyle = -16;
    private const int GwlExStyle = -20;
    private const int GwlWndProc = -4;
    private const int IddiApplication = 32512;
    private const long WsCaption = 0x00C00000L;
    private const long WsThickFrame = 0x00040000L;
    private const long WsBorder = 0x00800000L;
    private const long WsDlgFrame = 0x00400000L;
    private const long WsExDlgModalFrame = 0x00000001L;
    private const long WsExWindowEdge = 0x00000100L;
    private const long WsExClientEdge = 0x00000200L;
    private const uint NifMessage = 0x00000001;
    private const uint NifIcon = 0x00000002;
    private const uint NifTip = 0x00000004;
    private const uint NimAdd = 0x00000000;
    private const uint NimDelete = 0x00000002;
    private const uint MfString = 0x00000000;
    private const uint MfSeparator = 0x00000800;
    private const uint MfChecked = 0x00000008;
    private const uint TpmLeftAlign = 0x0000;
    private const uint TpmBottomAlign = 0x0020;
    private const uint TpmRightButton = 0x0002;
    private const uint MenuShowSidebar = 1001;
    private const uint MenuOpenWorkspace = 1002;
    private const uint MenuHide = 1003;
    private const uint MenuStartup = 1004;
    private const uint MenuExit = 1005;
    private const int DwmwaWindowCornerPreference = 33;
    private const int DwmwaBorderColor = 34;
    private static readonly uint DwmColorNone = 0xFFFFFFFE;
    private const uint DwmWindowCornerPreferenceDoNotRound = 1;
    private const int SwHide = 0;
    private const int SwShow = 5;
    private const int SwRestore = 9;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpFrameChanged = 0x0020;
    private readonly Microsoft.UI.Dispatching.DispatcherQueue _dispatcherQueue;
    private readonly WorkspaceSession _session;
    private readonly CanvasWindow _canvasWindow;
    private readonly IntPtr _hwnd;
    private AppWindow? _appWindow;
    private NotifyIconData _notifyIconData;
    private bool _trayIconRegistered;
    private bool _isExitRequested;
    private WindowProc? _windowProc;
    private nint _previousWindowProc;

    public MainWindow(WorkspaceSession session)
    {
        _session = session;
        _dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();

        InitializeComponent();
        RootGrid.DataContext = this;
        RootGrid.AddHandler(UIElement.KeyDownEvent, new KeyEventHandler(RootGrid_KeyDown), true);

        _hwnd = WindowNative.GetWindowHandle(this);
        _canvasWindow = new CanvasWindow(_session);
        _canvasWindow.SidebarRequested += CanvasWindow_SidebarRequested;

        ConfigureWindow();
        RegisterGlobalHotKey();
        InitializeTrayIcon();
        UpdateTaskSelection(null);

        Activated += MainWindow_Activated;
        Closed += MainWindow_Closed;
    }

    public ObservableCollection<TaskBoard> Tasks => _session?.Tasks ?? [];

    public void HideToBackground()
    {
        HideSidebar();
    }

    public void ShowSidebarFromProtocol()
    {
        ClearSidebarStatus();
        ShowSidebar();
    }

    public void ShowSidebarFromProtocol(string message, bool isWarning = false)
    {
        ShowSidebar();
        SetSidebarStatus(message, isWarning);
    }

    public bool TryOpenTask(Guid taskId)
    {
        var task = _session.Tasks.FirstOrDefault(candidate => candidate.Id == taskId);
        if (task is null)
        {
            return false;
        }

        TaskListView.SelectedItem = task;
        UpdateTaskSelection(task);
        OpenTask(task);
        return true;
    }

    private void ConfigureWindow()
    {
        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(_hwnd);
        _appWindow = AppWindow.GetFromWindowId(windowId);
        _appWindow.Closing += AppWindow_Closing;

        if (_appWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.IsMaximizable = false;
            presenter.IsMinimizable = false;
            presenter.IsResizable = false;
            presenter.SetBorderAndTitleBar(false, false);
        }

        RemoveNativeWindowFrame();
        RemoveSystemWindowCorner();
        RemoveSystemWindowBorder();

        PositionSidebarWindow();
    }

    private void RemoveNativeWindowFrame()
    {
        var style = GetWindowLongPtr(_hwnd, GwlStyle).ToInt64();
        style &= ~(WsCaption | WsThickFrame | WsBorder | WsDlgFrame);
        SetWindowLongPtr(_hwnd, GwlStyle, new nint(style));

        var exStyle = GetWindowLongPtr(_hwnd, GwlExStyle).ToInt64();
        exStyle &= ~(WsExDlgModalFrame | WsExWindowEdge | WsExClientEdge);
        SetWindowLongPtr(_hwnd, GwlExStyle, new nint(exStyle));

        SetWindowPos(_hwnd, IntPtr.Zero, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoZOrder | SwpFrameChanged);
    }

    private void RemoveSystemWindowBorder()
    {
        var borderColor = DwmColorNone;
        DwmSetWindowAttribute(_hwnd, DwmwaBorderColor, ref borderColor, Marshal.SizeOf<uint>());
    }

    private void RemoveSystemWindowCorner()
    {
        var cornerPreference = DwmWindowCornerPreferenceDoNotRound;
        DwmSetWindowAttribute(_hwnd, DwmwaWindowCornerPreference, ref cornerPreference, Marshal.SizeOf<uint>());
    }

    private void PositionSidebarWindow()
    {
        if (_appWindow is null)
        {
            return;
        }

        var workArea = DisplayArea.GetFromWindowId(_appWindow.Id, DisplayAreaFallback.Primary).WorkArea;
        _appWindow.MoveAndResize(new RectInt32(
            workArea.X + Math.Max(0, workArea.Width - 492),
            workArea.Y,
            492,
            Math.Max(680, workArea.Height)));
    }

    private void RegisterGlobalHotKey()
    {
        _windowProc = WindowMessageHandler;
        _previousWindowProc = SetWindowLongPtr(_hwnd, GwlWndProc, Marshal.GetFunctionPointerForDelegate(_windowProc));
        RegisterHotKey(_hwnd, HotKeyId, ModControl | ModNoRepeat, (uint)VirtualKey.Tab);
    }

    private void InitializeTrayIcon()
    {
        _notifyIconData = new NotifyIconData
        {
            cbSize = (uint)Marshal.SizeOf<NotifyIconData>(),
            hWnd = _hwnd,
            uID = 1,
            uFlags = NifMessage | NifIcon | NifTip,
            uCallbackMessage = WmTrayIcon,
            hIcon = LoadIcon(IntPtr.Zero, new IntPtr(IddiApplication)),
            szTip = "LumiCanvas"
        };

        _trayIconRegistered = ShellNotifyIcon(NimAdd, ref _notifyIconData);
    }

    private void AppWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_isExitRequested)
        {
            return;
        }

        args.Cancel = true;
        HideSidebar();
    }

    private void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
    {
        if (args.WindowActivationState == WindowActivationState.Deactivated)
        {
            HideSidebar();
        }
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        _isExitRequested = true;
        _session.Flush();
        UnregisterHotKey(_hwnd, HotKeyId);
        RemoveTrayIcon();
        _canvasWindow.Shutdown();

        if (_previousWindowProc != 0)
        {
            SetWindowLongPtr(_hwnd, GwlWndProc, _previousWindowProc);
        }
    }

    private nint WindowMessageHandler(nint hWnd, uint message, nint wParam, nint lParam)
    {
        if (message == WmHotKey && wParam.ToInt32() == HotKeyId)
        {
            _dispatcherQueue.TryEnqueue(() =>
            {
                if (IsVisible())
                {
                    HideSidebar();
                }
                else
                {
                    ShowSidebar();
                }
            });
            return 0;
        }

        if (message == WmTrayIcon)
        {
            var trayMessage = unchecked((uint)lParam.ToInt64());
            if (trayMessage == WmLButtonDblClk)
            {
                _dispatcherQueue.TryEnqueue(ShowSidebar);
                return 0;
            }

            if (trayMessage == WmRButtonUp)
            {
                _dispatcherQueue.TryEnqueue(ShowTrayMenu);
                return 0;
            }
        }

        if (message == WmCommand)
        {
            switch (LowWord(wParam))
            {
                case MenuShowSidebar:
                    _dispatcherQueue.TryEnqueue(ShowSidebar);
                    return 0;
                case MenuOpenWorkspace:
                    _dispatcherQueue.TryEnqueue(OpenCurrentWorkspace);
                    return 0;
                case MenuHide:
                    _dispatcherQueue.TryEnqueue(HideSidebar);
                    return 0;
                case MenuStartup:
                    _dispatcherQueue.TryEnqueue(ToggleStartupRegistration);
                    return 0;
                case MenuExit:
                    _dispatcherQueue.TryEnqueue(ExitApplication);
                    return 0;
            }
        }

        return CallWindowProc(_previousWindowProc, hWnd, message, wParam, lParam);
    }

    private void ShowSidebar()
    {
        PositionSidebarWindow();
        TaskListView.SelectedItem = _session.CurrentTask;
        UpdateCommandState();
        ShowWindow(_hwnd, SwShow);
        Activate();
        SetForegroundWindow(_hwnd);
    }

    private void HideSidebar()
    {
        ShowWindow(_hwnd, SwHide);
    }

    private bool IsVisible()
    {
        return IsWindowVisible(_hwnd);
    }

    private void CanvasWindow_SidebarRequested(object? sender, EventArgs e)
    {
        ShowSidebar();
    }

    private void AddTaskButton_Click(object sender, RoutedEventArgs e)
    {
        ClearSidebarStatus();
        AddTaskFromInput();
    }

    private void OpenArchiveFolderButton_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(_session.StorageFolderPath);

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/root,\"{_session.StorageFolderPath}\"",
            UseShellExecute = true
        });
    }

    private void CopyCurrentTaskLinkButton_Click(object sender, RoutedEventArgs e)
    {
        if (_session.CurrentTask is null)
        {
            SetSidebarStatus("当前没有可复制链接的任务。", true);
            return;
        }

        CopyTaskLinkToClipboard(_session.CurrentTask);
        SetSidebarStatus("已复制当前任务链接。", false);
    }

    private void RootGrid_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == VirtualKey.Escape)
        {
            HideSidebar();
            e.Handled = true;
        }
    }

    private void NewTaskTextBox_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == VirtualKey.Enter)
        {
            AddTaskFromInput();
            e.Handled = true;
        }
    }

    private void AddTaskFromInput()
    {
        var title = NewTaskTextBox.Text?.Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            return;
        }

        var task = _session.AddTask(title);
        NewTaskTextBox.Text = string.Empty;
        TaskListView.SelectedItem = task;
        UpdateTaskSelection(task);
        OpenTask(task);
    }

    private void TaskListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateTaskSelection(TaskListView.SelectedItem as TaskBoard);
    }

    private void TaskListView_ItemClick(object sender, ItemClickEventArgs e)
    {
        if (e.ClickedItem is not TaskBoard task)
        {
            return;
        }

        TaskListView.SelectedItem = task;
        UpdateTaskSelection(task);
        OpenTask(task);
    }

    private void UpdateTaskSelection(TaskBoard? task)
    {
        _session.CurrentTask = task;
        UpdateCommandState();
    }

    private void OpenTask(TaskBoard task)
    {
        ClearSidebarStatus();
        _session.CurrentTask = task;
        _canvasWindow.ShowTask(task);
        HideSidebar();
    }

    private void OpenCurrentWorkspace()
    {
        if (_session.CurrentTask is not null)
        {
            _canvasWindow.ShowTask(_session.CurrentTask);
            return;
        }

        ShowSidebar();
    }

    private void TaskItem_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not TaskBoard task)
        {
            return;
        }

        var flyout = new MenuFlyout();
        var completeItem = new MenuFlyoutItem { Text = "完成任务", Tag = task };
        completeItem.Click += CompleteTaskMenuItem_Click;
        flyout.Items.Add(completeItem);

        var abandonItem = new MenuFlyoutItem { Text = "放弃任务", Tag = task };
        abandonItem.Click += AbandonTaskMenuItem_Click;
        flyout.Items.Add(abandonItem);

        flyout.ShowAt(element);
        e.Handled = true;
    }

    private void CompleteTaskMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuFlyoutItem menuItem && menuItem.Tag is TaskBoard task)
        {
            task.State = TaskBoardState.Completed;
        }
    }

    private void AbandonTaskMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuFlyoutItem menuItem && menuItem.Tag is TaskBoard task)
        {
            task.State = TaskBoardState.Abandoned;
        }
    }

    private void UpdateCommandState()
    {
        CopyCurrentTaskLinkButton.IsEnabled = _session.CurrentTask is not null;
    }

    private static string BuildTaskProtocolLink(TaskBoard task)
    {
        return $"lumicanvas://open?taskId={task.Id}";
    }

    private static void CopyTaskLinkToClipboard(TaskBoard task)
    {
        var package = new DataPackage();
        package.SetText(BuildTaskProtocolLink(task));
        Clipboard.SetContent(package);
        Clipboard.Flush();
    }

    private void SetSidebarStatus(string message, bool isWarning = false)
    {
        SidebarStatusTextBlock.Text = message;
        SidebarStatusTextBlock.Foreground = isWarning
            ? new Microsoft.UI.Xaml.Media.SolidColorBrush(ColorHelper.FromArgb(255, 255, 196, 120))
            : new Microsoft.UI.Xaml.Media.SolidColorBrush(ColorHelper.FromArgb(255, 140, 196, 255));
        SidebarStatusTextBlock.Visibility = Visibility.Visible;
    }

    private void ClearSidebarStatus()
    {
        SidebarStatusTextBlock.Text = string.Empty;
        SidebarStatusTextBlock.Visibility = Visibility.Collapsed;
    }

    private void ToggleStartupRegistration()
    {
        SetStartupRegistration(!IsStartupRegistrationEnabled());
    }

    private static bool IsStartupRegistrationEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", writable: false);
        return key?.GetValue(StartupRegistryValueName) is string value && !string.IsNullOrWhiteSpace(value);
    }

    private static void SetStartupRegistration(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
        if (key is null)
        {
            return;
        }

        if (!enabled)
        {
            key.DeleteValue(StartupRegistryValueName, false);
            return;
        }

        var executablePath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(executablePath))
        {
            key.SetValue(StartupRegistryValueName, $"\"{executablePath}\"");
        }
    }

    private void ExitApplication()
    {
        _isExitRequested = true;
        _session.Flush();
        RemoveTrayIcon();
        _canvasWindow.Shutdown();
        Close();
        Application.Current.Exit();
    }

    private void ShowTrayMenu()
    {
        var menu = CreatePopupMenu();
        if (menu == IntPtr.Zero)
        {
            return;
        }

        try
        {
            AppendMenu(menu, MfString, MenuShowSidebar, "显示侧边栏");
            AppendMenu(menu, MfString, MenuOpenWorkspace, _session.CurrentTask is null ? "打开当前画布（无任务）" : "打开当前画布");
            AppendMenu(menu, MfString, MenuHide, "隐藏");
            AppendMenu(menu, MfSeparator, 0, string.Empty);
            AppendMenu(menu, MfString | (IsStartupRegistrationEnabled() ? MfChecked : 0), MenuStartup, "开机启动");
            AppendMenu(menu, MfSeparator, 0, string.Empty);
            AppendMenu(menu, MfString, MenuExit, "退出");

            GetCursorPos(out var cursorPosition);
            SetForegroundWindow(_hwnd);
            TrackPopupMenuEx(menu, TpmLeftAlign | TpmBottomAlign | TpmRightButton, cursorPosition.X, cursorPosition.Y, _hwnd, IntPtr.Zero);
            PostMessage(_hwnd, 0, IntPtr.Zero, IntPtr.Zero);
        }
        finally
        {
            DestroyMenu(menu);
        }
    }

    private void RemoveTrayIcon()
    {
        if (!_trayIconRegistered)
        {
            return;
        }

        ShellNotifyIcon(NimDelete, ref _notifyIconData);
        _trayIconRegistered = false;
    }

    private static uint LowWord(nint value)
    {
        return (uint)(value.ToInt64() & 0xFFFF);
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern nint SetWindowLongPtr(IntPtr hWnd, int nIndex, nint dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern nint GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "CallWindowProcW")]
    private static extern nint CallWindowProc(nint lpPrevWndFunc, nint hWnd, uint msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadIcon(IntPtr hInstance, IntPtr lpIconName);

    [DllImport("shell32.dll", EntryPoint = "Shell_NotifyIconW", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShellNotifyIcon(uint dwMessage, ref NotifyIconData lpData);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref uint pvAttribute, int cbAttribute);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreatePopupMenu();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AppendMenu(IntPtr hMenu, uint uFlags, uint uIDNewItem, string lpNewItem);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyMenu(IntPtr hMenu);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll")]
    private static extern int TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    private delegate nint WindowProc(nint hWnd, uint message, nint wParam, nint lParam);
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct NotifyIconData
{
    public uint cbSize;
    public IntPtr hWnd;
    public uint uID;
    public uint uFlags;
    public uint uCallbackMessage;
    public IntPtr hIcon;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
    public string szTip;
    public uint dwState;
    public uint dwStateMask;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
    public string szInfo;
    public uint uVersion;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
    public string szInfoTitle;
    public uint dwInfoFlags;
    public Guid guidItem;
    public IntPtr hBalloonIcon;
}

[StructLayout(LayoutKind.Sequential)]
public struct NativePoint
{
    public int X;
    public int Y;
}
