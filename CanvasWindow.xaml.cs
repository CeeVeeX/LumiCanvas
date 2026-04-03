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
    private const double MinimumTimeTagWidth = 240;
    private const double MinimumTimeTagHeight = 180;
    private const double MiniMapPaddingWorld = 120;
    private const double MiniMapCanvasPadding = 4;
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
    private bool _isDragDeferredSaveActive;
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
    private Point _lastCanvasPointerPosition;
    private bool _hasCanvasPointerPosition;
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
    private List<ClipboardBoardItemPayload> _copiedBoardItems = [];

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
        InitializeClipboardIntegration();
        CanvasViewport.SizeChanged += CanvasViewport_SizeChanged;

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
        PositionCanvasWindowToCursorDisplay();
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
        UninitializeClipboardIntegration();
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

        PositionCanvasWindowToCursorDisplay();

        HideWindow();
    }

    private void PositionCanvasWindowToCursorDisplay()
    {
        if (_appWindow is null)
        {
            return;
        }

        GetCursorPos(out var cursorPosition);
        var workArea = DisplayArea.GetFromPoint(new PointInt32(cursorPosition.X, cursorPosition.Y), DisplayAreaFallback.Primary).WorkArea;
        _appWindow.MoveAndResize(new RectInt32(
            workArea.X,
            workArea.Y,
            workArea.Width,
            workArea.Height));
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

    private async void RootGrid_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Escape)
        {
            ReturnToSidebar();
            e.Handled = true;
            return;
        }

        if (e.Key == Windows.System.VirtualKey.V && IsControlPressed())
        {
            if (!IsTextInputFocused() && await PasteFromClipboardAsync())
            {
                e.Handled = true;
                return;
            }
        }

        if (e.Key == Windows.System.VirtualKey.C && IsControlPressed())
        {
            if (!IsTextInputFocused() && await CopySelectedItemsToClipboardAsync())
            {
                e.Handled = true;
                return;
            }
        }

        if (e.Key == Windows.System.VirtualKey.X && IsControlPressed())
        {
            if (!IsTextInputFocused() && await CopySelectedItemsToClipboardAsync())
            {
                DeleteSelectedItems();
                e.Handled = true;
                return;
            }
        }

        if (e.Key == Windows.System.VirtualKey.Back)
        {
            DeleteSelectedItems();
            e.Handled = true;
        }
    }

    private void DeleteSelectedItems()
    {
        if (_session.CurrentTask is null || _selectedItemIds.Count == 0)
        {
            return;
        }

        var removingItems = _session.CurrentTask.Items
            .Where(item => _selectedItemIds.Contains(item.Id) && !item.IsEditing)
            .ToList();

        if (removingItems.Count == 0)
        {
            return;
        }

        _session.BeginDeferredSave();
        try
        {
            foreach (var item in removingItems)
            {
                _session.CurrentTask.Items.Remove(item);
                RemoveBoardItemView(item.Id);
                _selectedItemIds.Remove(item.Id);
            }
        }
        finally
        {
            _session.EndDeferredSave();
        }

        ClearSelectedItems();
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

    private void AddTimeTagMenuItem_Click(object sender, RoutedEventArgs e)
    {
        AddTimeTagAt(_contextMenuWorldPoint == default ? GetViewportCenterWorldPoint() : _contextMenuWorldPoint);
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

    private void AddTimeTagAt(Point position)
    {
        if (_session.CurrentTask is null)
        {
            return;
        }

        _session.CurrentTask.Items.Add(new BoardItemModel
        {
            Kind = BoardItemKind.TimeTag,
            X = position.X,
            Y = position.Y,
            Width = 320,
            Height = 220,
            ZIndex = _highestZIndex + 1,
            Content = "时间标签",
            TimeTagDueAt = DateTimeOffset.Now.AddHours(1),
            TimeTagReminderEnabled = true,
            TimeTagRecurrence = TimeTagRecurrence.None,
            TimeTagLastReminderAt = null
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

        var addTimeTagItem = new MenuFlyoutItem { Text = "添加时间标签" };
        addTimeTagItem.Click += AddTimeTagMenuItem_Click;
        flyout.Items.Add(addTimeTagItem);

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

    private Point GetPreferredPasteWorldPoint()
    {
        if (_appWindow is not null && GetCursorPos(out var cursorScreenPoint))
        {
            var windowPosition = _appWindow.Position;
            var rasterScale = Math.Max(0.0001, CanvasViewport.XamlRoot?.RasterizationScale ?? 1.0);
            var localX = (cursorScreenPoint.X - windowPosition.X) / rasterScale;
            var localY = (cursorScreenPoint.Y - windowPosition.Y) / rasterScale;
            var x = Math.Clamp(localX, 0, Math.Max(0, CanvasViewport.ActualWidth));
            var y = Math.Clamp(localY, 0, Math.Max(0, CanvasViewport.ActualHeight));
            return ScreenToWorld(new Point(x, y));
        }

        if (_hasCanvasPointerPosition)
        {
            var x = Math.Clamp(_lastCanvasPointerPosition.X, 0, Math.Max(0, CanvasViewport.ActualWidth));
            var y = Math.Clamp(_lastCanvasPointerPosition.Y, 0, Math.Max(0, CanvasViewport.ActualHeight));
            return ScreenToWorld(new Point(x, y));
        }

        return GetViewportCenterWorldPoint();
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
        UpdateMiniMap();
    }

    private void CanvasViewport_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateMiniMap();
    }

    private void UpdateMiniMap()
    {
        if (MiniMapCanvas is null || MiniMapContainer is null)
        {
            return;
        }

        if (_session.CurrentTask is null)
        {
            MiniMapCanvas.Children.Clear();
            MiniMapContainer.Visibility = Visibility.Collapsed;
            return;
        }

        MiniMapContainer.Visibility = Visibility.Visible;

        var mapWidth = Math.Max(1, MiniMapCanvas.Width - MiniMapCanvasPadding * 2);
        var mapHeight = Math.Max(1, MiniMapCanvas.Height - MiniMapCanvasPadding * 2);

        var visibleLeft = -_offsetX / _scale;
        var visibleTop = -_offsetY / _scale;
        var visibleWidth = Math.Max(1, CanvasViewport.ActualWidth / _scale);
        var visibleHeight = Math.Max(1, CanvasViewport.ActualHeight / _scale);
        var visibleRight = visibleLeft + visibleWidth;
        var visibleBottom = visibleTop + visibleHeight;

        var items = _session.CurrentTask.Items;
        var hasItems = items.Count > 0;

        var minX = hasItems ? items.Min(item => item.X) : visibleLeft;
        var minY = hasItems ? items.Min(item => item.Y) : visibleTop;
        var maxX = hasItems ? items.Max(item => item.X + item.Width) : visibleRight;
        var maxY = hasItems ? items.Max(item => item.Y + item.Height) : visibleBottom;

        minX = Math.Min(minX, visibleLeft) - MiniMapPaddingWorld;
        minY = Math.Min(minY, visibleTop) - MiniMapPaddingWorld;
        maxX = Math.Max(maxX, visibleRight) + MiniMapPaddingWorld;
        maxY = Math.Max(maxY, visibleBottom) + MiniMapPaddingWorld;

        var worldWidth = Math.Max(1, maxX - minX);
        var worldHeight = Math.Max(1, maxY - minY);
        var scale = Math.Min(mapWidth / worldWidth, mapHeight / worldHeight);

        var contentWidth = worldWidth * scale;
        var contentHeight = worldHeight * scale;
        var originX = MiniMapCanvasPadding + (mapWidth - contentWidth) / 2;
        var originY = MiniMapCanvasPadding + (mapHeight - contentHeight) / 2;

        MiniMapCanvas.Children.Clear();

        foreach (var item in items)
        {
            var rect = new Microsoft.UI.Xaml.Shapes.Rectangle
            {
                Width = Math.Max(2, item.Width * scale),
                Height = Math.Max(2, item.Height * scale),
                Fill = new SolidColorBrush(ColorHelper.FromArgb(120, 140, 190, 255)),
                Stroke = new SolidColorBrush(ColorHelper.FromArgb(180, 190, 225, 255)),
                StrokeThickness = 0.5,
                RadiusX = 1,
                RadiusY = 1,
                IsHitTestVisible = false
            };

            Canvas.SetLeft(rect, originX + (item.X - minX) * scale);
            Canvas.SetTop(rect, originY + (item.Y - minY) * scale);
            MiniMapCanvas.Children.Add(rect);
        }

        var viewportRect = new Microsoft.UI.Xaml.Shapes.Rectangle
        {
            Width = Math.Max(4, visibleWidth * scale),
            Height = Math.Max(4, visibleHeight * scale),
            Fill = new SolidColorBrush(ColorHelper.FromArgb(35, 120, 188, 255)),
            Stroke = new SolidColorBrush(ColorHelper.FromArgb(240, 120, 188, 255)),
            StrokeThickness = 1,
            IsHitTestVisible = false
        };

        Canvas.SetLeft(viewportRect, originX + (visibleLeft - minX) * scale);
        Canvas.SetTop(viewportRect, originY + (visibleTop - minY) * scale);
        MiniMapCanvas.Children.Add(viewportRect);
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
            AddBoardItemView(item);
        }

        _highestZIndex = _session.CurrentTask.Items.Count == 0 ? 0 : _session.CurrentTask.Items.Max(item => item.ZIndex);
        UpdateMiniMap();
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
        _isDragDeferredSaveActive = false;
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
                BringItemToFront(_draggedItem, element);
                if (!_isDragDeferredSaveActive)
                {
                    _session.BeginDeferredSave();
                    _isDragDeferredSaveActive = true;
                }
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

        UpdateMiniMap();

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
        if (_isDragDeferredSaveActive)
        {
            _session.EndDeferredSave();
            _isDragDeferredSaveActive = false;
        }
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
        UpdateMiniMap();
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

    private bool IsTextInputFocused()
    {
        var root = Content?.XamlRoot;
        if (root is null)
        {
            return false;
        }

        var focused = FocusManager.GetFocusedElement(root) as DependencyObject;
        while (focused is not null)
        {
            if (focused is TextBox or RichEditBox)
            {
                return true;
            }

            focused = VisualTreeHelper.GetParent(focused);
        }

        return false;
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
            BoardItemKind.TimeTag => (MinimumTimeTagWidth, MinimumTimeTagHeight),
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

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern nint SetWindowLongPtr(IntPtr hWnd, int nIndex, nint dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern nint GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref uint pvAttribute, int cbAttribute);
}
