using System;
using System.IO;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.Web.WebView2.Core;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private const string PdfViewDeleteMessage = "__lumi_delete_selected_pdf__";

    private FrameworkElement BuildPdfCard(BoardItemModel item)
    {
        if (string.IsNullOrWhiteSpace(item.SourcePath) || !File.Exists(item.SourcePath))
        {
            return CreateMissingMediaHint("PDF文件不存在或已被移动");
        }

        var layout = new Grid();
        layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var fileName = Path.GetFileName(item.SourcePath);
        var toolbar = new Grid
        {
            Padding = new Thickness(8),
            ColumnSpacing = 8,
            ColumnDefinitions =
            {
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                new ColumnDefinition { Width = GridLength.Auto }
            }
        };

        var fileNameText = new TextBlock
        {
            Text = fileName,
            FontSize = 13,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Foreground = ItemTitleBrush
        };

        var dragHandle = new Border
        {
            Padding = new Thickness(10, 0, 10, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = "拖动",
                FontSize = 12,
                Opacity = 0.65
            }
        };

        Grid.SetColumn(fileNameText, 0);
        Grid.SetColumn(dragHandle, 1);
        toolbar.Children.Add(fileNameText);
        toolbar.Children.Add(dragHandle);

        var webView = new WebView2
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            DefaultBackgroundColor = Windows.UI.Color.FromArgb(255, 82, 82, 82)
        };
        webView.PointerPressed += (_, args) =>
        {
            var point = args.GetCurrentPoint(CanvasViewport);
            if (!point.Properties.IsMiddleButtonPressed)
            {
                return;
            }

            if (!_selectedItemIds.Contains(item.Id))
            {
                SetSelectedItems([item]);
            }

            _isPanning = true;
            _isDraggingItem = false;
            _draggedItem = null;
            _activeDraggedItems.Clear();
            _pressedItem = item;
            _pressedItemElement = webView;
            _lastPointerPosition = point.Position;
            webView.CapturePointer(args.Pointer);
            args.Handled = true;
        };
        webView.PointerMoved += (_, args) =>
        {
            if (!_isPanning || !ReferenceEquals(_pressedItemElement, webView))
            {
                return;
            }

            var currentPosition = args.GetCurrentPoint(CanvasViewport).Position;
            _offsetX += currentPosition.X - _lastPointerPosition.X;
            _offsetY += currentPosition.Y - _lastPointerPosition.Y;
            _lastPointerPosition = currentPosition;
            UpdateCanvasTransform();
            args.Handled = true;
        };
        webView.PointerReleased += (_, args) =>
        {
            if (!ReferenceEquals(_pressedItemElement, webView))
            {
                return;
            }

            webView.ReleasePointerCapture(args.Pointer);
            _isPanning = false;
            _pressedItem = null;
            _pressedItemElement = null;
            args.Handled = true;
        };

        webView.GotFocus += (_, _) =>
        {
            if (!_selectedItemIds.Contains(item.Id))
            {
                SetSelectedItems([item]);
            }
        };

        Grid.SetRow(toolbar, 0);
        Grid.SetRow(webView, 1);
        layout.Children.Add(toolbar);
        layout.Children.Add(webView);

        async void InitializePdfViewer(object? _, RoutedEventArgs __)
        {
            webView.Loaded -= InitializePdfViewer;
            try
            {
                var environment = await GetWebView2EnvironmentAsync();
                await webView.EnsureCoreWebView2Async(environment);

                webView.CoreWebView2.WebMessageReceived += (_, args) =>
                {
                    if (!string.Equals(args.TryGetWebMessageAsString(), PdfViewDeleteMessage, StringComparison.Ordinal))
                    {
                        return;
                    }

                    _dispatcherQueue.TryEnqueue(() =>
                    {
                        if (!_selectedItemIds.Contains(item.Id))
                        {
                            SetSelectedItems([item]);
                        }

                        DeleteSelectedItems();
                    });
                };

                await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(
                    """
                    (() => {
                        if (window.__lumiDeleteHookInstalled) {
                            return;
                        }

                        window.__lumiDeleteHookInstalled = true;
                        window.addEventListener('keydown', (event) => {
                            const key = event.key;
                            if (key !== 'Delete' && key !== 'Backspace') {
                                return;
                            }

                            const target = event.target;
                            const tagName = target && target.tagName ? target.tagName.toUpperCase() : '';
                            const isEditable = !!(target && (target.isContentEditable || tagName === 'INPUT' || tagName === 'TEXTAREA' || tagName === 'SELECT'));
                            if (isEditable) {
                                return;
                            }

                            if (window.chrome && window.chrome.webview) {
                                window.chrome.webview.postMessage('__lumi_delete_selected_pdf__');
                                event.preventDefault();
                                event.stopPropagation();
                            }
                        }, true);
                    })();
                    """);

                var absolutePath = Path.GetFullPath(item.SourcePath);
                var fileUri = new Uri(absolutePath).AbsoluteUri;
                webView.CoreWebView2.Navigate(fileUri);
            }
            catch
            {
                // 如果WebView2初始化失败，显示错误消息
            }
        }

        webView.Loaded += InitializePdfViewer;

        return layout;
    }
}
