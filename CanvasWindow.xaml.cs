using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.UI.Input;
using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.Web.WebView2.Core;
using DocText = Microsoft.UI.Text;
using FontText = Windows.UI.Text;
using Windows.Foundation;
using Windows.Graphics;
using Windows.Media.Core;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace LumiCanvas;

public sealed partial class CanvasWindow : Window
{
    private const double DragStartThreshold = 6;
    private const double ResizeHandleSize = 18;
    private const double MinimumMarkdownWidth = 240;
    private const double MinimumMarkdownHeight = 140;
    private const double MinimumMediaWidth = 220;
    private const double MinimumMediaHeight = 160;
    private const double MinimumFileWidth = 220;
    private const double MinimumFileHeight = 100;
    private const int GwlStyle = -16;
    private const int GwlExStyle = -20;
    private const int DwmwaWindowCornerPreference = 33;
    private const int DwmwaBorderColor = 34;
    private static readonly uint DwmColorNone = 0xFFFFFFFE;
    private const uint DwmWindowCornerPreferenceDoNotRound = 1;
    private const long WsCaption = 0x00C00000L;
    private const long WsThickFrame = 0x00040000L;
    private const long WsBorder = 0x00800000L;
    private const long WsDlgFrame = 0x00400000L;
    private const long WsExDlgModalFrame = 0x00000001L;
    private const long WsExWindowEdge = 0x00000100L;
    private const long WsExClientEdge = 0x00000200L;
    private const int SwHide = 0;
    private const int SwRestore = 9;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpFrameChanged = 0x0020;
    private static readonly FontText.FontWeight NormalWeight = new() { Weight = 400 };
    private static readonly FontText.FontWeight SemiBoldWeight = new() { Weight = 600 };
    private static readonly FontText.FontWeight BoldWeight = new() { Weight = 700 };
    private static readonly SolidColorBrush ItemBackgroundBrush = new(ColorHelper.FromArgb(255, 26, 32, 41));
    private static readonly SolidColorBrush ItemBorderBrush = new(ColorHelper.FromArgb(255, 52, 66, 85));
    private static readonly SolidColorBrush ItemTitleBrush = new(Colors.White);
    private static readonly SolidColorBrush ItemTextBrush = new(ColorHelper.FromArgb(255, 214, 224, 236));
    private static readonly SolidColorBrush SecondaryTextBrush = new(ColorHelper.FromArgb(255, 152, 168, 186));
    private static readonly SolidColorBrush AccentBrush = new(ColorHelper.FromArgb(255, 104, 180, 255));
    private readonly DispatcherQueue _dispatcherQueue;
    private readonly DispatcherQueueTimer _markdownHighlightTimer;
    private readonly WorkspaceSession _session;
    private readonly IntPtr _hwnd;
    private readonly Dictionary<Guid, FrameworkElement> _itemViews = [];
    private readonly HashSet<Guid> _selectedItemIds = [];
    private AppWindow? _appWindow;
    private bool _isPanning;
    private bool _isDraggingItem;
    private bool _isResizingItem;
    private bool _isSelectingArea;
    private bool _isApplyingMarkdownHighlight;
    private bool _isInternalClose;
    private double _scale = 1;
    private double _offsetX;
    private double _offsetY;
    private Point _lastPointerPosition;
    private Point _pressedPointerPosition;
    private Point _resizeStartPointerPosition;
    private Point _contextMenuWorldPoint;
    private Point _selectionStartPoint;
    private BoardItemModel? _draggedItem;
    private BoardItemModel? _resizedItem;
    private FrameworkElement? _pressedItemElement;
    private FrameworkElement? _resizeItemElement;
    private BoardItemModel? _pressedItem;
    private Guid? _pendingFocusItemId;
    private RichEditBox? _pendingHighlightEditor;
    private int _highestZIndex;
    private double _resizeStartWidth;
    private double _resizeStartHeight;
    private List<BoardItemModel> _activeDraggedItems = [];

    public CanvasWindow(WorkspaceSession session)
    {
        _session = session;
        _dispatcherQueue = DispatcherQueue.GetForCurrentThread();
        _markdownHighlightTimer = _dispatcherQueue.CreateTimer();
        _markdownHighlightTimer.IsRepeating = false;
        _markdownHighlightTimer.Interval = TimeSpan.FromMilliseconds(180);
        _markdownHighlightTimer.Tick += MarkdownHighlightTimer_Tick;

        InitializeComponent();
        _hwnd = WindowNative.GetWindowHandle(this);
        RootGrid.AddHandler(UIElement.KeyDownEvent, new KeyEventHandler(RootGrid_KeyDown), true);

        ConfigureWindow();
        UpdateCommandState();
        UpdateCanvasTransform();

        Activated += CanvasWindow_Activated;
    }

    public event EventHandler? SidebarRequested;

    public void ShowTask(TaskBoard task)
    {
        _session.CurrentTask = task;
        CurrentTaskTitleText.Text = task.Title;
        UpdateCommandState();
        RenderCurrentBoard();
        ShowWindow(_hwnd, SwRestore);
        Activate();
        SetForegroundWindow(_hwnd);
    }

    public void HideWindow()
    {
        ShowWindow(_hwnd, SwHide);
    }

    public void Shutdown()
    {
        _isInternalClose = true;
        Close();
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

        var workArea = DisplayArea.GetFromWindowId(windowId, DisplayAreaFallback.Primary).WorkArea;
        _appWindow.MoveAndResize(new RectInt32(
            workArea.X,
            workArea.Y,
            workArea.Width,
            workArea.Height));

        HideWindow();
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

    private void AppWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_isInternalClose)
        {
            return;
        }

        args.Cancel = true;
        HideWindow();
    }

    private void CanvasWindow_Activated(object sender, WindowActivatedEventArgs args)
    {
        if (_session.CurrentTask is null)
        {
            UpdateCommandState();
        }
    }

    private void BackToSidebarButton_Click(object sender, RoutedEventArgs e)
    {
        ReturnToSidebar();
    }

    private void CloseCanvasButton_Click(object sender, RoutedEventArgs e)
    {
        ReturnToSidebar();
    }

    private void RootGrid_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Escape)
        {
            ReturnToSidebar();
            e.Handled = true;
        }
    }

    private void ReturnToSidebar()
    {
        HideWindow();
        SidebarRequested?.Invoke(this, EventArgs.Empty);
    }

    private void UpdateCommandState()
    {
        EmptyBoardHintCard.Visibility = _session.CurrentTask is null ? Visibility.Visible : Visibility.Collapsed;
        CurrentTaskTitleText.Text = _session.CurrentTask?.Title ?? "未选择任务";
    }

    private async void AddFileMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await AddFileAtAsync(_contextMenuWorldPoint == default ? GetViewportCenterWorldPoint() : _contextMenuWorldPoint);
    }

    private async void AddFolderMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await AddFolderAtAsync(_contextMenuWorldPoint == default ? GetViewportCenterWorldPoint() : _contextMenuWorldPoint);
    }

    private async Task AddFileAtAsync(Point position)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        var picker = new FileOpenPicker();
        InitializeWithWindow.Initialize(picker, _hwnd);
        picker.FileTypeFilter.Add("*");

        var file = await picker.PickSingleFileAsync();
        if (file is null || string.IsNullOrWhiteSpace(file.Path))
        {
            return;
        }

        var kind = ResolveFileKind(file.Path);

        _session.CurrentTask.Items.Add(new BoardItemModel
        {
            Kind = kind,
            X = position.X,
            Y = position.Y,
            Width = kind == BoardItemKind.Image ? 360 : kind == BoardItemKind.Video ? 420 : 320,
            Height = kind == BoardItemKind.Image ? 240 : kind == BoardItemKind.Video ? 300 : 140,
            ZIndex = _highestZIndex + 1,
            SourcePath = file.Path
        });

        RenderCurrentBoard();
    }

    private async Task AddFolderAtAsync(Point position)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        var picker = new FolderPicker();
        InitializeWithWindow.Initialize(picker, _hwnd);
        picker.FileTypeFilter.Add("*");

        var folder = await picker.PickSingleFolderAsync();
        if (folder is null || string.IsNullOrWhiteSpace(folder.Path))
        {
            return;
        }

        _session.CurrentTask.Items.Add(new BoardItemModel
        {
            Kind = BoardItemKind.File,
            X = position.X,
            Y = position.Y,
            Width = 320,
            Height = 140,
            ZIndex = _highestZIndex + 1,
            SourcePath = folder.Path
        });

        RenderCurrentBoard();
    }

    private void CanvasViewport_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        var clickedItem = FindBoardItemFromSource(e.OriginalSource as DependencyObject);
        if (ExitEditingIfNeeded(clickedItem))
        {
            e.Handled = true;
            return;
        }

        if (clickedItem is not null)
        {
            return;
        }

        var point = e.GetCurrentPoint(CanvasViewport);
        if (!point.Properties.IsLeftButtonPressed)
        {
            return;
        }

        if (IsControlPressed())
        {
            _isSelectingArea = true;
            _selectionStartPoint = point.Position;
            UpdateSelectionMarquee(point.Position);
            CanvasViewport.CapturePointer(e.Pointer);
            e.Handled = true;
            return;
        }

        ClearSelectedItems();

        _isPanning = true;
        _lastPointerPosition = point.Position;
        CanvasViewport.CapturePointer(e.Pointer);
        e.Handled = true;
    }

    private void CanvasViewport_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (_isSelectingArea)
        {
            UpdateSelectionMarquee(e.GetCurrentPoint(CanvasViewport).Position);
            e.Handled = true;
            return;
        }

        if (!_isPanning)
        {
            return;
        }

        var currentPosition = e.GetCurrentPoint(CanvasViewport).Position;
        _offsetX += currentPosition.X - _lastPointerPosition.X;
        _offsetY += currentPosition.Y - _lastPointerPosition.Y;
        _lastPointerPosition = currentPosition;
        UpdateCanvasTransform();
        e.Handled = true;
    }

    private void CanvasViewport_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (_isSelectingArea)
        {
            _isSelectingArea = false;
            CanvasViewport.ReleasePointerCapture(e.Pointer);
            ApplySelectionMarquee(e.GetCurrentPoint(CanvasViewport).Position);
            SelectionMarquee.Visibility = Visibility.Collapsed;
            e.Handled = true;
            return;
        }

        _isPanning = false;
        CanvasViewport.ReleasePointerCapture(e.Pointer);
    }

    private void CanvasViewport_PointerWheelChanged(object sender, PointerRoutedEventArgs e)
    {
        var pointerPoint = e.GetCurrentPoint(CanvasViewport);
        var zoomFactor = pointerPoint.Properties.MouseWheelDelta > 0 ? 1.12 : 1 / 1.12;
        var worldBeforeZoom = ScreenToWorld(pointerPoint.Position);
        _scale = Math.Clamp(_scale * zoomFactor, 0.25, 4.5);
        _offsetX = pointerPoint.Position.X - (worldBeforeZoom.X * _scale);
        _offsetY = pointerPoint.Position.Y - (worldBeforeZoom.Y * _scale);
        UpdateCanvasTransform();
        e.Handled = true;
    }

    private void CanvasViewport_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (_session.CurrentTask is null || FindBoardItemFromSource(e.OriginalSource as DependencyObject) is not null)
        {
            return;
        }

        _session.CurrentTask.Items.Add(CreateMarkdownItem(ScreenToWorld(e.GetPosition(CanvasViewport))));
        RenderCurrentBoard();
        e.Handled = true;
    }

    private void CanvasViewport_RightTapped(object sender, RightTappedRoutedEventArgs e)
    {
        if (_session.CurrentTask is null || FindBoardItemFromSource(e.OriginalSource as DependencyObject) is not null)
        {
            return;
        }

        _contextMenuWorldPoint = ScreenToWorld(e.GetPosition(CanvasViewport));

        var flyout = new MenuFlyout();
        var addFileItem = new MenuFlyoutItem { Text = "添加文件" };
        addFileItem.Click += AddFileMenuItem_Click;
        flyout.Items.Add(addFileItem);

        var addFolderItem = new MenuFlyoutItem { Text = "添加文件夹" };
        addFolderItem.Click += AddFolderMenuItem_Click;
        flyout.Items.Add(addFolderItem);

        flyout.ShowAt(CanvasViewport, e.GetPosition(CanvasViewport));
        e.Handled = true;
    }

    private BoardItemModel CreateMarkdownItem(Point position)
    {
        return new BoardItemModel
        {
            Kind = BoardItemKind.Markdown,
            X = position.X,
            Y = position.Y,
            Width = 380,
            Height = 260,
            ZIndex = _highestZIndex + 1,
            IsEditing = false,
            Content = "# 新笔记\n\n双击卡片进入编辑模式。"
        };
    }

    private Point GetViewportCenterWorldPoint()
    {
        return ScreenToWorld(new Point(Math.Max(180, CanvasViewport.ActualWidth / 2), Math.Max(120, CanvasViewport.ActualHeight / 2)));
    }

    private Point ScreenToWorld(Point screenPoint)
    {
        return new Point((screenPoint.X - _offsetX) / _scale, (screenPoint.Y - _offsetY) / _scale);
    }

    private void UpdateCanvasTransform()
    {
        BoardTransform.ScaleX = _scale;
        BoardTransform.ScaleY = _scale;
        BoardTransform.TranslateX = _offsetX;
        BoardTransform.TranslateY = _offsetY;
    }

    private void RenderCurrentBoard()
    {
        BoardCanvas.Children.Clear();
        _itemViews.Clear();
        UpdateCommandState();

        if (_session.CurrentTask is null)
        {
            return;
        }

        foreach (var item in _session.CurrentTask.Items)
        {
            var view = CreateBoardItemView(item);
            _itemViews[item.Id] = view;
            Canvas.SetLeft(view, item.X);
            Canvas.SetTop(view, item.Y);
            Canvas.SetZIndex(view, item.ZIndex);
            BoardCanvas.Children.Add(view);
        }

        _highestZIndex = _session.CurrentTask.Items.Count == 0 ? 0 : _session.CurrentTask.Items.Max(item => item.ZIndex);
    }

    private FrameworkElement CreateBoardItemView(BoardItemModel item)
    {
        var shell = new Border
        {
            Tag = item,
            Width = item.Width,
            Height = item.Height,
            CornerRadius = new CornerRadius(0),
            BorderThickness = new Thickness(0),
            Background = item.Kind == BoardItemKind.Markdown ? new SolidColorBrush(Colors.Transparent) : ItemBackgroundBrush,
            Shadow = item.Kind == BoardItemKind.Markdown ? null : new ThemeShadow(),
            Translation = item.Kind == BoardItemKind.Markdown ? new Vector3(0, 0, 0) : new Vector3(0, 6, 24)
        };

        shell.PointerPressed += BoardItem_PointerPressed;
        shell.PointerMoved += BoardItem_PointerMoved;
        shell.PointerReleased += BoardItem_PointerReleased;
        shell.PointerEntered += BoardItem_PointerEntered;
        shell.PointerExited += BoardItem_PointerExited;
        shell.DoubleTapped += BoardItem_DoubleTapped;

        var content = item.Kind switch
        {
            BoardItemKind.Image => BuildImageCard(item),
            BoardItemKind.Video => BuildVideoCard(item),
            BoardItemKind.File => BuildFileCard(item),
            _ => BuildMarkdownCard(item)
        };

        var resizeHandle = BuildResizeHandle(item);
        var layout = new Grid();

        layout.Children.Add(new Microsoft.UI.Xaml.Shapes.Rectangle
        {
            Tag = "SelectionOutline",
            Visibility = _selectedItemIds.Contains(item.Id) ? Visibility.Visible : Visibility.Collapsed,
            Stroke = new SolidColorBrush(ColorHelper.FromArgb(200, 140, 196, 255)),
            StrokeThickness = 1,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            IsHitTestVisible = false
        });

        if (item.Kind == BoardItemKind.Markdown)
        {
            layout.Children.Add(new Microsoft.UI.Xaml.Shapes.Rectangle
            {
                Tag = "ResizeOutline",
                Visibility = Visibility.Collapsed,
                Stroke = new SolidColorBrush(ColorHelper.FromArgb(220, 160, 200, 255)),
                StrokeThickness = 1,
                StrokeDashArray = [4, 4],
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                IsHitTestVisible = false
            });
        }

        layout.Children.Add(content);
        layout.Children.Add(resizeHandle);

        shell.Child = layout;

        return shell;
    }

    private FrameworkElement BuildResizeHandle(BoardItemModel item)
    {
        var iconBrush = new SolidColorBrush(ColorHelper.FromArgb(200, 18, 24, 32));

        var handle = new Border
        {
            DataContext = item,
            Tag = "ResizeHandle",
            Width = ResizeHandleSize,
            Height = ResizeHandleSize,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
            Visibility = Visibility.Collapsed,
            Background = new SolidColorBrush(ColorHelper.FromArgb(96, 255, 255, 255)),
            Child = new Grid
            {
                Width = 10,
                Height = 10,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false,
                Children =
                {
                    new Border
                    {
                        Width = 1,
                        Height = 8,
                        Background = iconBrush,
                        HorizontalAlignment = HorizontalAlignment.Right,
                        VerticalAlignment = VerticalAlignment.Bottom
                    },
                    new Border
                    {
                        Width = 8,
                        Height = 1,
                        Background = iconBrush,
                        HorizontalAlignment = HorizontalAlignment.Right,
                        VerticalAlignment = VerticalAlignment.Bottom
                    },
                    new Border
                    {
                        Width = 1,
                        Height = 5,
                        Background = iconBrush,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Top
                    },
                    new Border
                    {
                        Width = 5,
                        Height = 1,
                        Background = iconBrush,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Top
                    }
                }
            }
        };

        handle.PointerPressed += ResizeHandle_PointerPressed;
        handle.PointerMoved += ResizeHandle_PointerMoved;
        handle.PointerReleased += ResizeHandle_PointerReleased;
        return handle;
    }

    private FrameworkElement BuildMarkdownCard(BoardItemModel item)
    {
        var layout = new Grid();

        if (item.IsEditing)
        {
            var editor = new WebView2
            {
                Tag = item,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };

            editor.WebMessageReceived += MarkdownWebView_WebMessageReceived;
            editor.Loaded += MarkdownEditor_Loaded;

            layout.Children.Add(editor);
            return layout;
        }

        layout.Children.Add(new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Padding = new Thickness(12),
            Content = BuildMarkdownPreview(item.Content ?? string.Empty)
        });
        return layout;
    }

    private FrameworkElement BuildImageCard(BoardItemModel item)
    {
        return !string.IsNullOrWhiteSpace(item.SourcePath) && File.Exists(item.SourcePath)
            ? new Image
            {
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Source = new BitmapImage(new Uri(item.SourcePath))
            }
            : CreateMissingMediaHint("图片文件不存在或已被移动");
    }

    private FrameworkElement BuildVideoCard(BoardItemModel item)
    {
        return !string.IsNullOrWhiteSpace(item.SourcePath) && File.Exists(item.SourcePath)
            ? BuildVideoPlayer(item)
            : CreateMissingMediaHint("视频文件不存在或已被移动");
    }

    private FrameworkElement BuildVideoPlayer(BoardItemModel item)
    {
        var mediaPlayer = new MediaPlayerElement
        {
            Stretch = Stretch.Uniform,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            AreTransportControlsEnabled = false,
            Source = MediaSource.CreateFromUri(new Uri(item.SourcePath!))
        };

        mediaPlayer.PointerEntered += VideoPlayer_PointerEntered;
        mediaPlayer.PointerExited += VideoPlayer_PointerExited;
        mediaPlayer.PointerCanceled += VideoPlayer_PointerExited;
        mediaPlayer.PointerCaptureLost += VideoPlayer_PointerCaptureLost;
        return mediaPlayer;
    }

    private FrameworkElement BuildFileCard(BoardItemModel item)
    {
        var layout = new Border
        {
            Background = new SolidColorBrush(ColorHelper.FromArgb(200, 14, 20, 28)),
            Padding = new Thickness(12),
            Child = new StackPanel
            {
                Spacing = 8,
                VerticalAlignment = VerticalAlignment.Center,
                Children =
                {
                    new TextBlock
                    {
                        Foreground = SecondaryTextBrush,
                        TextWrapping = TextWrapping.WrapWholeWords,
                        Text = IsDirectoryItem(item) ? "点击名称打开该文件夹" : "点击名称打开所在文件夹并选中该文件"
                    }
                }
            }
        };

        if (layout.Child is not StackPanel stackPanel)
        {
            return layout;
        }

        var openButton = new Button
        {
            Tag = item,
            HorizontalAlignment = HorizontalAlignment.Left,
            Padding = new Thickness(0),
            Background = new SolidColorBrush(Colors.Transparent),
            BorderThickness = new Thickness(0),
            Content = new TextBlock
            {
                Foreground = AccentBrush,
                TextWrapping = TextWrapping.WrapWholeWords,
                Text = GetItemDisplayName(item)
            }
        };
        openButton.Click += FileItemOpenButton_Click;
        stackPanel.Children.Add(openButton);
        return layout;
    }

    private FrameworkElement CreateMissingMediaHint(string message)
    {
        return new Border
        {
            CornerRadius = new CornerRadius(0),
            Background = new SolidColorBrush(ColorHelper.FromArgb(255, 14, 20, 28)),
            Child = new TextBlock
            {
                Margin = new Thickness(16),
                Foreground = SecondaryTextBrush,
                TextWrapping = TextWrapping.Wrap,
                Text = message
            }
        };
    }

    private FrameworkElement BuildMarkdownPreview(string markdown)
    {
        var panel = new StackPanel { Spacing = 8 };
        var lines = (markdown ?? string.Empty).Replace("\r\n", "\n").Split('\n');
        var codeBuffer = string.Empty;
        var insideCodeBlock = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine ?? string.Empty;
            if (line.StartsWith("```", StringComparison.Ordinal))
            {
                if (insideCodeBlock)
                {
                    panel.Children.Add(CreateCodeBlock(codeBuffer.TrimEnd('\n')));
                    codeBuffer = string.Empty;
                }

                insideCodeBlock = !insideCodeBlock;
                continue;
            }

            if (insideCodeBlock)
            {
                codeBuffer += line + "\n";
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                panel.Children.Add(new Microsoft.UI.Xaml.Shapes.Rectangle { Height = 2, Opacity = 0 });
                continue;
            }

            if (TryCreateMarkdownImage(line, out var imageElement))
            {
                panel.Children.Add(imageElement);
                continue;
            }

            panel.Children.Add(CreateMarkdownLine(line));
        }

        if (!string.IsNullOrWhiteSpace(codeBuffer))
        {
            panel.Children.Add(CreateCodeBlock(codeBuffer.TrimEnd('\n')));
        }

        if (panel.Children.Count == 0)
        {
            panel.Children.Add(new TextBlock
            {
                Foreground = SecondaryTextBrush,
                Text = "空白笔记",
                FontStyle = FontText.FontStyle.Italic
            });
        }

        return panel;
    }

    private FrameworkElement CreateMarkdownImage(string source, string altText)
    {
        if (Uri.TryCreate(source, UriKind.Absolute, out var uri))
        {
            return CreateImageElementWithFallback(new BitmapImage(uri), source, altText);
        }

        if (File.Exists(source))
        {
            return CreateImageElementWithFallback(new BitmapImage(new Uri(source)), source, altText);
        }

        return CreateStyledTextBlock(string.IsNullOrWhiteSpace(altText)
            ? $"[图片加载失败] {source}"
            : $"[图片加载失败] {altText} ({source})", 13, NormalWeight);
    }

    private FrameworkElement CreateImageElementWithFallback(ImageSource source, string rawSource, string altText)
    {
        var fallback = CreateStyledTextBlock(string.IsNullOrWhiteSpace(altText)
            ? $"[图片加载失败] {rawSource}"
            : $"[图片加载失败] {altText} ({rawSource})", 13, NormalWeight);
        fallback.Visibility = Visibility.Collapsed;

        var image = new Image
        {
            Source = source,
            Stretch = Stretch.Uniform,
            MaxHeight = 320,
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Top
        };

        image.ImageFailed += (_, _) =>
        {
            image.Visibility = Visibility.Collapsed;
            fallback.Visibility = Visibility.Visible;
        };

        var container = new Grid();
        container.Children.Add(image);
        container.Children.Add(fallback);
        return container;
    }

    private bool TryCreateMarkdownImage(string line, out FrameworkElement imageElement)
    {
        var linkedImageMatch = Regex.Match(line, "^\\[!\\[(.*?)\\]\\((.*?)\\)\\]\\((.*?)\\)$");
        if (linkedImageMatch.Success)
        {
            var linkedAltText = linkedImageMatch.Groups[1].Value.Trim();
            var linkedImageSource = linkedImageMatch.Groups[2].Value.Trim();
            var linkTarget = linkedImageMatch.Groups[3].Value.Trim();

            if (string.IsNullOrWhiteSpace(linkedImageSource))
            {
                imageElement = CreateStyledTextBlock("[图片地址为空]", 13, NormalWeight);
                return true;
            }

            var renderedImage = CreateMarkdownImage(linkedImageSource, linkedAltText);

            if (!string.IsNullOrWhiteSpace(linkTarget) &&
                (Uri.TryCreate(linkTarget, UriKind.Absolute, out var navigateUri) ||
                 (File.Exists(linkTarget) && Uri.TryCreate(linkTarget, UriKind.Absolute, out navigateUri))))
            {
                imageElement = new HyperlinkButton
                {
                    NavigateUri = navigateUri,
                    Content = renderedImage,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Padding = new Thickness(0),
                    BorderThickness = new Thickness(0),
                    Background = new SolidColorBrush(Colors.Transparent)
                };
                return true;
            }

            imageElement = renderedImage;
            return true;
        }

        var match = Regex.Match(line, "^!\\[(.*?)\\]\\((.*?)\\)$");
        if (!match.Success)
        {
            imageElement = null!;
            return false;
        }

        var altText = match.Groups[1].Value.Trim();
        var source = match.Groups[2].Value.Trim();
        if (string.IsNullOrWhiteSpace(source))
        {
            imageElement = CreateStyledTextBlock("[图片地址为空]", 13, NormalWeight);
            return true;
        }

        imageElement = CreateMarkdownImage(source, altText);
        return true;
    }

    private FrameworkElement CreateMarkdownLine(string line)
    {
        if (line.StartsWith("### ", StringComparison.Ordinal))
        {
            return CreateStyledTextBlock(line[4..], 18, SemiBoldWeight);
        }

        if (line.StartsWith("## ", StringComparison.Ordinal))
        {
            return CreateStyledTextBlock(line[3..], 22, SemiBoldWeight);
        }

        if (line.StartsWith("# ", StringComparison.Ordinal))
        {
            return CreateStyledTextBlock(line[2..], 28, BoldWeight);
        }

        if (line.StartsWith("- ", StringComparison.Ordinal) || line.StartsWith("* ", StringComparison.Ordinal))
        {
            var bulletLine = CreateStyledTextBlock($"\u2022 {line[2..]}", 14, NormalWeight);
            bulletLine.FontFamily = new FontFamily("Segoe UI Symbol");
            return bulletLine;
        }

        if (line.StartsWith("> ", StringComparison.Ordinal))
        {
            var quote = CreateStyledTextBlock(line[2..], 14, NormalWeight);
            quote.Foreground = AccentBrush;
            return quote;
        }

        return CreateStyledTextBlock(line, 14, NormalWeight);
    }

    private TextBlock CreateStyledTextBlock(string text, double fontSize, FontText.FontWeight fontWeight)
    {
        var textBlock = new TextBlock
        {
            FontSize = fontSize,
            FontWeight = fontWeight,
            Foreground = ItemTextBrush,
            TextWrapping = TextWrapping.WrapWholeWords,
            LineHeight = fontSize + 8
        };

        AddMarkdownInlines(textBlock.Inlines, text);
        return textBlock;
    }

    private Border CreateCodeBlock(string code)
    {
        return new Border
        {
            Padding = new Thickness(12),
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(ColorHelper.FromArgb(255, 14, 20, 28)),
            Child = new TextBlock
            {
                FontFamily = new FontFamily("Consolas"),
                Foreground = new SolidColorBrush(ColorHelper.FromArgb(255, 181, 225, 168)),
                TextWrapping = TextWrapping.Wrap,
                Text = code
            }
        };
    }

    private static BoardItemKind ResolveFileKind(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" => BoardItemKind.Image,
            ".mp4" or ".mov" or ".wmv" or ".avi" or ".mkv" or ".webm" => BoardItemKind.Video,
            _ => BoardItemKind.File
        };
    }

    private static string GetItemDisplayName(BoardItemModel item)
    {
        if (item.Kind == BoardItemKind.Markdown)
        {
            return item.Content ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(item.SourcePath))
        {
            return "未命名文件";
        }

        if (Directory.Exists(item.SourcePath))
        {
            var normalizedPath = Path.TrimEndingDirectorySeparator(item.SourcePath);
            var folderName = Path.GetFileName(normalizedPath);
            return string.IsNullOrWhiteSpace(folderName) ? normalizedPath : folderName;
        }

        return Path.GetFileName(item.SourcePath);
    }

    private void FileItemOpenButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not BoardItemModel item || string.IsNullOrWhiteSpace(item.SourcePath))
        {
            return;
        }

        if (!File.Exists(item.SourcePath) && !Directory.Exists(item.SourcePath))
        {
            return;
        }

        if (Directory.Exists(item.SourcePath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = item.SourcePath,
                UseShellExecute = true
            });
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,\"{item.SourcePath}\"",
            UseShellExecute = true
        });
    }

    private static bool IsDirectoryItem(BoardItemModel item)
    {
        return !string.IsNullOrWhiteSpace(item.SourcePath) && Directory.Exists(item.SourcePath);
    }

    private void AddMarkdownInlines(InlineCollection inlines, string text)
    {
        var index = 0;
        while (index < text.Length)
        {
            var boldIndex = text.IndexOf("**", index, StringComparison.Ordinal);
            var codeIndex = text.IndexOf('`', index);
            var linkIndex = text.IndexOf('[', index);
            while (linkIndex > 0 && text[linkIndex - 1] == '!')
            {
                linkIndex = text.IndexOf('[', linkIndex + 1);
            }

            var nextIndex = MinPositive(MinPositive(boldIndex, codeIndex), linkIndex);

            if (nextIndex < 0)
            {
                inlines.Add(new Run { Text = text[index..] });
                return;
            }

            if (nextIndex > index)
            {
                inlines.Add(new Run { Text = text[index..nextIndex] });
            }

            if (nextIndex == boldIndex)
            {
                var end = text.IndexOf("**", boldIndex + 2, StringComparison.Ordinal);
                if (end > boldIndex)
                {
                    inlines.Add(new Run
                    {
                        Text = text[(boldIndex + 2)..end],
                        FontWeight = SemiBoldWeight,
                        Foreground = AccentBrush
                    });
                    index = end + 2;
                    continue;
                }
            }

            if (nextIndex == codeIndex)
            {
                var end = text.IndexOf('`', codeIndex + 1);
                if (end > codeIndex)
                {
                    inlines.Add(new Run
                    {
                        Text = text[(codeIndex + 1)..end],
                        FontFamily = new FontFamily("Consolas"),
                        Foreground = new SolidColorBrush(ColorHelper.FromArgb(255, 255, 196, 120))
                    });
                    index = end + 1;
                    continue;
                }
            }

            if (nextIndex == linkIndex)
            {
                var textEnd = text.IndexOf(']', linkIndex + 1);
                if (textEnd > linkIndex && textEnd + 1 < text.Length && text[textEnd + 1] == '(')
                {
                    var linkEnd = text.IndexOf(')', textEnd + 2);
                    if (linkEnd > textEnd + 2)
                    {
                        var linkText = text[(linkIndex + 1)..textEnd];
                        var linkTarget = text[(textEnd + 2)..linkEnd];
                        if (Uri.TryCreate(linkTarget, UriKind.Absolute, out var navigateUri))
                        {
                            var hyperlink = new Hyperlink();
                            hyperlink.Click += (_, _) => HandleMarkdownHyperlink(navigateUri);
                            hyperlink.Inlines.Add(new Run
                            {
                                Text = linkText,
                                Foreground = AccentBrush
                            });
                            inlines.Add(hyperlink);
                        }
                        else
                        {
                            inlines.Add(new Run { Text = text[linkIndex..(linkEnd + 1)] });
                        }

                        index = linkEnd + 1;
                        continue;
                    }
                }
            }

            inlines.Add(new Run { Text = text[nextIndex].ToString() });
            index = nextIndex + 1;
        }
    }

    private void HandleMarkdownHyperlink(Uri uri)
    {
        if (TryOpenTaskFromProtocolUri(uri))
        {
            return;
        }

        _ = Windows.System.Launcher.LaunchUriAsync(uri);
    }

    private bool TryOpenTaskFromProtocolUri(Uri uri)
    {
        if (!string.Equals(uri.Scheme, "lumicanvas", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!TryGetTaskIdFromUri(uri, out var taskId))
        {
            return false;
        }

        var task = _session.Tasks.FirstOrDefault(candidate => candidate.Id == taskId);
        if (task is null)
        {
            return false;
        }

        ShowTask(task);
        return true;
    }

    private static bool TryGetTaskIdFromUri(Uri uri, out Guid taskId)
    {
        taskId = Guid.Empty;

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

    private static int MinPositive(int first, int second)
    {
        if (first < 0)
        {
            return second;
        }

        if (second < 0)
        {
            return first;
        }

        return Math.Min(first, second);
    }

    private void MarkdownDoneButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is BoardItemModel item)
        {
            _markdownHighlightTimer.Stop();
            _pendingHighlightEditor = null;
            item.IsEditing = false;
            RenderCurrentBoard();
        }
    }

    private async void MarkdownEditor_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not WebView2 editor || editor.Tag is not BoardItemModel item)
        {
            return;
        }

        async void OnNavigationCompleted(WebView2 webView, CoreWebView2NavigationCompletedEventArgs _)
        {
            webView.NavigationCompleted -= OnNavigationCompleted;
            var markdownJson = JsonSerializer.Serialize(item.Content ?? string.Empty);
            await webView.ExecuteScriptAsync($"window.setMarkdownValue({markdownJson});");

            if (_pendingFocusItemId == item.Id)
            {
                _pendingFocusItemId = null;
                await webView.ExecuteScriptAsync("window.focusMarkdownEditor();");
            }
        }

        await editor.EnsureCoreWebView2Async();

        var monacoLoaderUrl = ConfigureMonacoLocalAssets(editor);
        var monacoVsBaseUrl = string.IsNullOrWhiteSpace(monacoLoaderUrl)
            ? null
            : monacoLoaderUrl.Replace("/loader.js", string.Empty, StringComparison.Ordinal);
        editor.NavigationCompleted += OnNavigationCompleted;
        editor.NavigateToString(BuildMonacoEditorHtml(monacoLoaderUrl, monacoVsBaseUrl));
    }

    private string? ConfigureMonacoLocalAssets(WebView2 editor)
    {
        if (editor.CoreWebView2 is null)
        {
            return null;
        }

        var webAssetsPath = Path.Combine(AppContext.BaseDirectory, "WebAssets");
        var monacoPath = Path.Combine(webAssetsPath, "Monaco", "min", "vs", "loader.js");
        if (!File.Exists(monacoPath))
        {
            return null;
        }

        editor.CoreWebView2.SetVirtualHostNameToFolderMapping(
            "appassets",
            webAssetsPath,
            CoreWebView2HostResourceAccessKind.Allow);

        return "https://appassets/Monaco/min/vs/loader.js";
    }

    private void MarkdownWebView_WebMessageReceived(WebView2 sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        if (sender.Tag is not BoardItemModel item)
        {
            return;
        }

        item.Content = NormalizeEditorText(args.TryGetWebMessageAsString());
    }

    private static string BuildMonacoEditorHtml(string? monacoLoaderUrl, string? monacoVsBaseUrl)
    {
        var loaderScriptTag = string.IsNullOrWhiteSpace(monacoLoaderUrl)
            ? string.Empty
            : $"<script src=\"{monacoLoaderUrl}\"></script>";
        var monacoVsBaseUrlJson = JsonSerializer.Serialize(monacoVsBaseUrl ?? string.Empty);

        var html = """
                   <!doctype html>
                   <html>
                   <head>
                     <meta charset="utf-8" />
                     <style>
                       html, body, #container {
                         margin: 0;
                         width: 100%;
                         height: 100%;
                         background: transparent;
                         overflow: hidden;
                       }
                       #fallback {
                         width: 100%;
                         height: 100%;
                         border: 0;
                         outline: none;
                         resize: none;
                         box-sizing: border-box;
                         padding: 10px;
                         background: transparent;
                         color: #d6e0ec;
                         font: 14px Consolas, "Microsoft YaHei UI", sans-serif;
                       }
                     </style>
                     __LOADER_SCRIPT__
                   </head>
                   <body>
                     <div id="container"></div>
                     <textarea id="fallback" spellcheck="false"></textarea>
                     <script>
                       const fallback = document.getElementById('fallback');
                       let editor = null;
                       const monacoVsBase = __MONACO_VS_BASE_JSON__;

                       function postContent(value) {
                         if (window.chrome && window.chrome.webview) {
                           window.chrome.webview.postMessage(value || '');
                         }
                       }

                       window.setMarkdownValue = function (value) {
                         const text = value || '';
                         if (editor) {
                           editor.setValue(text);
                           return;
                         }
                         fallback.value = text;
                       };

                       window.focusMarkdownEditor = function () {
                         if (editor) {
                           editor.focus();
                           return;
                         }
                         fallback.focus();
                       };

                       fallback.addEventListener('input', () => postContent(fallback.value));

                       if (typeof require !== 'function' || !monacoVsBase) {
                         fallback.style.display = 'block';
                       } else {
                         require.config({
                           paths: {
                             vs: monacoVsBase
                           }
                         });

                         require(['vs/editor/editor.main'], function () {
                           fallback.style.display = 'none';
                           editor = monaco.editor.create(document.getElementById('container'), {
                             value: fallback.value,
                             language: 'markdown',
                             theme: 'vs-dark',
                             automaticLayout: true,
                             minimap: { enabled: false },
                             wordWrap: 'on',
                             lineNumbers: 'on',
                             renderLineHighlight: 'line',
                             scrollBeyondLastLine: false,
                             fontSize: 14
                           });

                           editor.onDidChangeModelContent(() => postContent(editor.getValue()));
                         });
                       }
                     </script>
                   </body>
                   </html>
                   """;

        return html
            .Replace("__LOADER_SCRIPT__", loaderScriptTag, StringComparison.Ordinal)
            .Replace("__MONACO_VS_BASE_JSON__", monacoVsBaseUrlJson, StringComparison.Ordinal);
    }

    private void MarkdownEditor_TextChanged(object sender, RoutedEventArgs e)
    {
        if (_isApplyingMarkdownHighlight || sender is not RichEditBox editor || editor.Tag is not BoardItemModel item)
        {
            return;
        }

        editor.Document.GetText(DocText.TextGetOptions.None, out var rawText);
        item.Content = NormalizeEditorText(rawText);
        _pendingHighlightEditor = editor;
        _markdownHighlightTimer.Stop();
        _markdownHighlightTimer.Start();
    }

    private void MarkdownHighlightTimer_Tick(DispatcherQueueTimer sender, object args)
    {
        if (_pendingHighlightEditor is null || _pendingHighlightEditor.XamlRoot is null)
        {
            return;
        }

        _pendingHighlightEditor.Document.GetText(DocText.TextGetOptions.None, out var rawText);
        ApplyMarkdownSyntaxHighlighting(_pendingHighlightEditor, rawText);
    }

    private void ApplyMarkdownSyntaxHighlighting(RichEditBox editor, string sourceText)
    {
        _isApplyingMarkdownHighlight = true;
        try
        {
            var selection = editor.Document.Selection;
            var start = selection.StartPosition;
            var end = selection.EndPosition;
            var fullRange = editor.Document.GetRange(0, sourceText.Length);
            fullRange.CharacterFormat.ForegroundColor = Colors.White;
            fullRange.CharacterFormat.Bold = DocText.FormatEffect.Off;
            fullRange.CharacterFormat.Italic = DocText.FormatEffect.Off;
            fullRange.CharacterFormat.BackgroundColor = Colors.Transparent;

            ApplyPattern(editor, sourceText, "^#{1,6}\\s.*$", Colors.DeepSkyBlue, DocText.FormatEffect.On, null, RegexOptions.Multiline);
            ApplyPattern(editor, sourceText, "^[-*]\\s.*$", Colors.Plum, DocText.FormatEffect.Off, null, RegexOptions.Multiline);
            ApplyPattern(editor, sourceText, "^>.*$", Colors.LightSteelBlue, DocText.FormatEffect.Off, DocText.FormatEffect.On, RegexOptions.Multiline);
            ApplyPattern(editor, sourceText, "```[\\s\\S]*?```", Colors.LightGreen, DocText.FormatEffect.Off, null, RegexOptions.Multiline);
            ApplyPattern(editor, sourceText, "\\*\\*.*?\\*\\*", Colors.LightGoldenrodYellow, DocText.FormatEffect.On, null, RegexOptions.None);
            ApplyPattern(editor, sourceText, "`[^`\\r\\n]+`", Colors.Orange, DocText.FormatEffect.Off, null, RegexOptions.None);

            editor.Document.Selection.SetRange(start, end);
        }
        finally
        {
            _isApplyingMarkdownHighlight = false;
        }
    }

    private void ApplyPattern(RichEditBox editor, string sourceText, string pattern, Windows.UI.Color color, DocText.FormatEffect bold, DocText.FormatEffect? italic, RegexOptions options)
    {
        foreach (Match match in Regex.Matches(sourceText, pattern, options))
        {
            var range = editor.Document.GetRange(match.Index, match.Index + match.Length);
            range.CharacterFormat.ForegroundColor = color;
            range.CharacterFormat.Bold = bold;
            if (italic.HasValue)
            {
                range.CharacterFormat.Italic = italic.Value;
            }
        }
    }

    private static string NormalizeEditorText(string rawText)
    {
        return rawText.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
    }

    private void BoardItem_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not BoardItemModel item)
        {
            return;
        }

        SetSelectedItems([item]);
        BringItemToFront(item, element);

        if (item.Kind == BoardItemKind.Markdown)
        {
            SetEditingItem(item);
            item.IsEditing = true;
            RenderCurrentBoard();
            e.Handled = true;
        }
    }

    private void BoardItem_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (_isResizingItem || sender is not FrameworkElement element || element.Tag is not BoardItemModel item || item.IsEditing)
        {
            return;
        }

        if (_selectedItemIds.Contains(item.Id))
        {
            EnsureSelectionOutlineVisible(item.Id);
        }
        else
        {
            SetSelectedItems([item]);
        }

        BringItemToFront(item, element);

        var pointerPoint = e.GetCurrentPoint(CanvasViewport);
        if (!pointerPoint.Properties.IsLeftButtonPressed || IsInteractiveElement(e.OriginalSource as DependencyObject))
        {
            return;
        }

        _isPanning = false;
        _isDraggingItem = false;
        _draggedItem = null;
        _activeDraggedItems = GetSelectedItemsForDrag(item);
        _pressedItem = item;
        _pressedItemElement = element;
        _pressedPointerPosition = pointerPoint.Position;
        _lastPointerPosition = pointerPoint.Position;
        _session.BeginDeferredSave();
        element.CapturePointer(e.Pointer);
        e.Handled = true;
    }

    private void BoardItem_PointerEntered(object sender, PointerRoutedEventArgs e)
    {
        if (sender is FrameworkElement element)
        {
            SetResizeHandleVisibility(element, true);
        }
    }

    private void BoardItem_PointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (sender is not FrameworkElement element)
        {
            return;
        }

        if (_isResizingItem && ReferenceEquals(element, _resizeItemElement))
        {
            return;
        }

        SetResizeHandleVisibility(element, false);
    }

    private void BoardItem_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (_isResizingItem)
        {
            return;
        }

        if (sender is not FrameworkElement element)
        {
            return;
        }

        var currentPosition = e.GetCurrentPoint(CanvasViewport).Position;

        if (!_isDraggingItem && _pressedItem is not null && ReferenceEquals(_pressedItemElement, element))
        {
            var deltaFromPressX = currentPosition.X - _pressedPointerPosition.X;
            var deltaFromPressY = currentPosition.Y - _pressedPointerPosition.Y;
            if (Math.Abs(deltaFromPressX) >= DragStartThreshold || Math.Abs(deltaFromPressY) >= DragStartThreshold)
            {
                _isDraggingItem = true;
                _draggedItem = _pressedItem;
                _lastPointerPosition = currentPosition;
            }
        }

        if (!_isDraggingItem || _draggedItem is null || !ReferenceEquals(element.Tag, _draggedItem))
        {
            return;
        }

        var deltaX = (currentPosition.X - _lastPointerPosition.X) / _scale;
        var deltaY = (currentPosition.Y - _lastPointerPosition.Y) / _scale;

        foreach (var draggedItem in _activeDraggedItems)
        {
            draggedItem.X += deltaX;
            draggedItem.Y += deltaY;

            if (_itemViews.TryGetValue(draggedItem.Id, out var draggedView))
            {
                Canvas.SetLeft(draggedView, draggedItem.X);
                Canvas.SetTop(draggedView, draggedItem.Y);
            }
        }

        _lastPointerPosition = currentPosition;
        e.Handled = true;
    }

    private void BoardItem_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (sender is UIElement element)
        {
            element.ReleasePointerCapture(e.Pointer);
        }

        _isDraggingItem = false;
        _draggedItem = null;
        _activeDraggedItems.Clear();
        _pressedItem = null;
        _pressedItemElement = null;
        _session.EndDeferredSave();
        if (sender is FrameworkElement shellElement)
        {
            SetResizeHandleVisibility(shellElement, true);
        }
        e.Handled = true;
    }

    private void ResizeHandle_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (sender is not FrameworkElement handle || handle.DataContext is not BoardItemModel item)
        {
            return;
        }

        var shell = FindItemShell(handle);
        if (shell is null)
        {
            return;
        }

        BringItemToFront(item, shell);
        _isDraggingItem = false;
        _draggedItem = null;
        _activeDraggedItems.Clear();
        _pressedItem = null;
        _pressedItemElement = null;
        _isResizingItem = true;
        _resizedItem = item;
        _resizeItemElement = shell;
        _resizeStartPointerPosition = e.GetCurrentPoint(CanvasViewport).Position;
        _resizeStartWidth = item.Width;
        _resizeStartHeight = item.Height;
        _session.BeginDeferredSave();
        SetSelectedItems([item]);
        SetResizeHandleVisibility(shell, true);
        SetResizeOutlineVisibility(_resizeItemElement, true);
        handle.CapturePointer(e.Pointer);
        e.Handled = true;
    }

    private void ResizeHandle_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isResizingItem || _resizedItem is null || _resizeItemElement is null || sender is not UIElement handle)
        {
            return;
        }

        var currentPosition = e.GetCurrentPoint(CanvasViewport).Position;
        var deltaX = (currentPosition.X - _resizeStartPointerPosition.X) / _scale;
        var deltaY = (currentPosition.Y - _resizeStartPointerPosition.Y) / _scale;
        var (minimumWidth, minimumHeight) = GetMinimumSize(_resizedItem.Kind);

        _resizedItem.Width = Math.Max(minimumWidth, _resizeStartWidth + deltaX);
        _resizedItem.Height = Math.Max(minimumHeight, _resizeStartHeight + deltaY);
        _resizeItemElement.Width = _resizedItem.Width;
        _resizeItemElement.Height = _resizedItem.Height;
        e.Handled = true;
    }

    private void ResizeHandle_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (sender is UIElement handle)
        {
            handle.ReleasePointerCapture(e.Pointer);
        }

        _isResizingItem = false;
        SetResizeOutlineVisibility(_resizeItemElement, false);
        SetResizeHandleVisibility(_resizeItemElement, true);
        _resizedItem = null;
        _resizeItemElement = null;
        _session.EndDeferredSave();
        e.Handled = true;
    }

    private BoardItemModel? FindBoardItemFromSource(DependencyObject? source)
    {
        while (source is not null)
        {
            if (source is FrameworkElement element && element.Tag is BoardItemModel item)
            {
                return item;
            }

            source = VisualTreeHelper.GetParent(source);
        }

        return null;
    }

    private static FrameworkElement? FindItemShell(DependencyObject? source)
    {
        source = VisualTreeHelper.GetParent(source);
        while (source is not null)
        {
            if (source is FrameworkElement element && element.Tag is BoardItemModel)
            {
                return element;
            }

            source = VisualTreeHelper.GetParent(source);
        }

        return null;
    }

    private static void SetResizeOutlineVisibility(FrameworkElement? shell, bool isVisible)
    {
        if (shell is not Border { Child: Grid layout })
        {
            return;
        }

        var outline = layout.Children
            .OfType<Microsoft.UI.Xaml.Shapes.Rectangle>()
            .FirstOrDefault(rectangle => string.Equals(rectangle.Tag as string, "ResizeOutline", StringComparison.Ordinal));

        if (outline is not null)
        {
            outline.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private static void SetResizeHandleVisibility(FrameworkElement? shell, bool isVisible)
    {
        if (shell is not Border { Child: Grid layout })
        {
            return;
        }

        var handle = layout.Children
            .OfType<FrameworkElement>()
            .FirstOrDefault(element => string.Equals(element.Tag as string, "ResizeHandle", StringComparison.Ordinal));

        if (handle is not null)
        {
            handle.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private static void SetSelectionOutlineVisibility(FrameworkElement? shell, bool isVisible)
    {
        if (shell is not Border { Child: Grid layout })
        {
            return;
        }

        var outline = layout.Children
            .OfType<Microsoft.UI.Xaml.Shapes.Rectangle>()
            .FirstOrDefault(rectangle => string.Equals(rectangle.Tag as string, "SelectionOutline", StringComparison.Ordinal));

        if (outline is not null)
        {
            outline.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void SetSelectedItems(IEnumerable<BoardItemModel> items)
    {
        foreach (var selectedId in _selectedItemIds.ToList())
        {
            if (_itemViews.TryGetValue(selectedId, out var previous))
            {
                SetSelectionOutlineVisibility(previous, false);
                SetResizeHandleVisibility(previous, false);
            }
        }

        _selectedItemIds.Clear();

        foreach (var item in items)
        {
            _selectedItemIds.Add(item.Id);
            EnsureSelectionOutlineVisible(item.Id);
        }
    }

    private void ClearSelectedItems()
    {
        if (_selectedItemIds.Count == 0)
        {
            return;
        }

        foreach (var selectedId in _selectedItemIds)
        {
            if (_itemViews.TryGetValue(selectedId, out var previous))
            {
                SetSelectionOutlineVisibility(previous, false);
                SetResizeHandleVisibility(previous, false);
            }
        }

        _selectedItemIds.Clear();
    }

    private void EnsureSelectionOutlineVisible(Guid itemId)
    {
        if (_itemViews.TryGetValue(itemId, out var shell))
        {
            SetSelectionOutlineVisibility(shell, true);
        }
    }

    private List<BoardItemModel> GetSelectedItemsForDrag(BoardItemModel anchorItem)
    {
        if (_session.CurrentTask is null)
        {
            return [anchorItem];
        }

        var selectedItems = _session.CurrentTask.Items
            .Where(item => _selectedItemIds.Contains(item.Id))
            .ToList();

        return selectedItems.Count == 0 ? [anchorItem] : selectedItems;
    }

    private void UpdateSelectionMarquee(Point currentPoint)
    {
        var x = Math.Min(_selectionStartPoint.X, currentPoint.X);
        var y = Math.Min(_selectionStartPoint.Y, currentPoint.Y);
        var width = Math.Abs(currentPoint.X - _selectionStartPoint.X);
        var height = Math.Abs(currentPoint.Y - _selectionStartPoint.Y);

        SelectionMarquee.Visibility = Visibility.Visible;
        SelectionMarquee.Width = width;
        SelectionMarquee.Height = height;
        SelectionMarquee.Margin = new Thickness(x, y, 0, 0);
    }

    private void ApplySelectionMarquee(Point currentPoint)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        var worldStart = ScreenToWorld(_selectionStartPoint);
        var worldEnd = ScreenToWorld(currentPoint);
        var selectionRect = new Rect(
            Math.Min(worldStart.X, worldEnd.X),
            Math.Min(worldStart.Y, worldEnd.Y),
            Math.Abs(worldEnd.X - worldStart.X),
            Math.Abs(worldEnd.Y - worldStart.Y));

        var selectedItems = _session.CurrentTask.Items
            .Where(item => Intersects(selectionRect, new Rect(item.X, item.Y, item.Width, item.Height)))
            .ToList();

        SetSelectedItems(selectedItems);
    }

    private static bool Intersects(Rect first, Rect second)
    {
        return first.X <= second.X + second.Width &&
               first.X + first.Width >= second.X &&
               first.Y <= second.Y + second.Height &&
               first.Y + first.Height >= second.Y;
    }

    private static bool IsControlPressed()
    {
        return InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
    }

    private bool ExitEditingIfNeeded(BoardItemModel? clickedItem)
    {
        if (_session.CurrentTask is null)
        {
            return false;
        }

        var editingItem = _session.CurrentTask.Items.FirstOrDefault(item => item.IsEditing);
        if (editingItem is null)
        {
            return false;
        }

        if (ReferenceEquals(editingItem, clickedItem))
        {
            return false;
        }

        editingItem.IsEditing = false;
        _markdownHighlightTimer.Stop();
        _pendingHighlightEditor = null;
        _pendingFocusItemId = null;
        RenderCurrentBoard();
        return true;
    }

    private void SetEditingItem(BoardItemModel editingItem)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        foreach (var item in _session.CurrentTask.Items)
        {
            item.IsEditing = ReferenceEquals(item, editingItem);
        }

        _pendingFocusItemId = editingItem.Id;
    }

    private static (double Width, double Height) GetMinimumSize(BoardItemKind kind)
    {
        return kind switch
        {
            BoardItemKind.Markdown => (MinimumMarkdownWidth, MinimumMarkdownHeight),
            BoardItemKind.Image or BoardItemKind.Video => (MinimumMediaWidth, MinimumMediaHeight),
            _ => (MinimumFileWidth, MinimumFileHeight)
        };
    }

    private void BringItemToFront(BoardItemModel item, FrameworkElement? element = null)
    {
        if (item.ZIndex <= _highestZIndex)
        {
            item.ZIndex = ++_highestZIndex;
        }

        if (element is not null)
        {
            Canvas.SetZIndex(element, item.ZIndex);
        }
    }

    private static bool IsInteractiveElement(DependencyObject? source)
    {
        while (source is not null)
        {
            if (source is RichEditBox || source is WebView2 || source is Button || source is HyperlinkButton || source is Slider)
            {
                return true;
            }

            source = VisualTreeHelper.GetParent(source);
        }

        return false;
    }

    private static void VideoPlayer_PointerEntered(object sender, PointerRoutedEventArgs e)
    {
        if (sender is MediaPlayerElement mediaPlayer)
        {
            mediaPlayer.AreTransportControlsEnabled = true;
        }
    }

    private static void VideoPlayer_PointerExited(object sender, PointerRoutedEventArgs e)
    {
        if (sender is MediaPlayerElement mediaPlayer)
        {
            mediaPlayer.AreTransportControlsEnabled = false;
        }
    }

    private static void VideoPlayer_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        if (sender is MediaPlayerElement mediaPlayer)
        {
            mediaPlayer.AreTransportControlsEnabled = false;
        }
    }

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern nint SetWindowLongPtr(IntPtr hWnd, int nIndex, nint dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern nint GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref uint pvAttribute, int cbAttribute);
}
