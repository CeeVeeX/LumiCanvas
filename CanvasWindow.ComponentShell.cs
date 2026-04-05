using System;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
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
            Translation = item.Kind == BoardItemKind.Markdown ? new System.Numerics.Vector3(0, 0, 0) : new System.Numerics.Vector3(0, 6, 24)
        };

        shell.PointerPressed += BoardItem_PointerPressed;
        shell.PointerMoved += BoardItem_PointerMoved;
        shell.PointerReleased += BoardItem_PointerReleased;
        shell.PointerEntered += BoardItem_PointerEntered;
        shell.PointerExited += BoardItem_PointerExited;
        shell.DoubleTapped += BoardItem_DoubleTapped;
        shell.RightTapped += BoardItem_RightTapped;

        var content = item.Kind switch
        {
            BoardItemKind.Image => BuildImageCard(item),
            BoardItemKind.Video => BuildVideoCard(item),
            BoardItemKind.File => BuildFileCard(item),
            BoardItemKind.TimeTag => BuildTimeTagCard(item),
            BoardItemKind.WebView => BuildWebViewCard(item),
            BoardItemKind.Pdf => BuildPdfCard(item),
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

    private void RefreshBoardItemView(BoardItemModel item)
    {
        if (!_itemViews.TryGetValue(item.Id, out var oldView))
        {
            RenderCurrentBoard();
            return;
        }

        var left = Canvas.GetLeft(oldView);
        var top = Canvas.GetTop(oldView);
        var zIndex = Canvas.GetZIndex(oldView);
        var index = BoardCanvas.Children.IndexOf(oldView);

        var newView = CreateBoardItemView(item);
        Canvas.SetLeft(newView, left);
        Canvas.SetTop(newView, top);
        Canvas.SetZIndex(newView, zIndex);

        if (index >= 0)
        {
            BoardCanvas.Children.RemoveAt(index);
            BoardCanvas.Children.Insert(index, newView);
        }
        else
        {
            BoardCanvas.Children.Remove(oldView);
            BoardCanvas.Children.Add(newView);
        }

        _itemViews[item.Id] = newView;
    }

    private FrameworkElement AddBoardItemView(BoardItemModel item)
    {
        var view = CreateBoardItemView(item);
        _itemViews[item.Id] = view;
        Canvas.SetLeft(view, item.X);
        Canvas.SetTop(view, item.Y);
        Canvas.SetZIndex(view, item.ZIndex);
        BoardCanvas.Children.Add(view);
        _highestZIndex = Math.Max(_highestZIndex, item.ZIndex);
        SetupFileWatcher(item);
        UpdateMiniMap();
        UpdateCanvasMetrics();
        return view;
    }

    private void RemoveBoardItemView(Guid itemId)
    {
        if (!_itemViews.TryGetValue(itemId, out var view))
        {
            return;
        }

        CleanupFileWatcher(itemId);
        BoardCanvas.Children.Remove(view);
        _itemViews.Remove(itemId);
        UpdateMiniMap();
        UpdateCanvasMetrics();
    }
}
