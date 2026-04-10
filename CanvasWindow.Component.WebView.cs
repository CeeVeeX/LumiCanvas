using System;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.Web.WebView2.Core;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private const string WebViewDeleteMessage = "__lumi_delete_selected__";

    private FrameworkElement BuildWebViewCard(BoardItemModel item)
    {
        var layout = new Grid();
        layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var backButton = new Button
        {
            Content = new FontIcon { Glyph = "\uE72B" },
            Width = 32,
            Height = 32,
            IsEnabled = false
        };

        var forwardButton = new Button
        {
            Content = new FontIcon { Glyph = "\uE72A" },
            Width = 32,
            Height = 32,
            IsEnabled = false
        };

        var refreshButton = new Button
        {
            Content = new FontIcon { Glyph = "\uE72C" },
            Width = 32,
            Height = 32
        };

        var addressBox = new TextBox
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            PlaceholderText = "输入网址并回车"
        };

        var toolbar = new Grid
        {
            Padding = new Thickness(8),
            ColumnSpacing = 8,
            ColumnDefinitions =
            {
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = GridLength.Auto },
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                new ColumnDefinition { Width = GridLength.Auto }
            }
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

        Grid.SetColumn(backButton, 0);
        Grid.SetColumn(forwardButton, 1);
        Grid.SetColumn(refreshButton, 2);
        Grid.SetColumn(addressBox, 3);
        Grid.SetColumn(dragHandle, 4);
        toolbar.Children.Add(backButton);
        toolbar.Children.Add(forwardButton);
        toolbar.Children.Add(refreshButton);
        toolbar.Children.Add(addressBox);
        toolbar.Children.Add(dragHandle);

        var webView = new WebView2
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            DefaultBackgroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255)
        };
        webView.PointerPressed += (_, args) =>
        {
            var point = args.GetCurrentPoint(CanvasViewport);
            if (!IsPanGesture(point))
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

        var initialAddress = string.IsNullOrWhiteSpace(item.Content) ? "https://www.bing.com" : item.Content!;
        addressBox.Text = initialAddress;

        async void InitializeWebView(object? _, RoutedEventArgs __)
        {
            webView.Loaded -= InitializeWebView;
            var environment = await GetWebView2EnvironmentAsync();
            await webView.EnsureCoreWebView2Async(environment);

            webView.CoreWebView2.WebMessageReceived += (_, args) =>
            {
                if (!string.Equals(args.TryGetWebMessageAsString(), WebViewDeleteMessage, StringComparison.Ordinal))
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
                            window.chrome.webview.postMessage('__lumi_delete_selected__');
                            event.preventDefault();
                            event.stopPropagation();
                        }
                    }, true);
                })();
                """);

            void UpdateNavButtonState()
            {
                backButton.IsEnabled = webView.CanGoBack;
                forwardButton.IsEnabled = webView.CanGoForward;
            }

            webView.CoreWebView2.HistoryChanged += (_, _) => UpdateNavButtonState();
            webView.NavigationCompleted += (_, _) =>
            {
                var current = webView.Source?.ToString() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(current))
                {
                    item.Content = current;
                    if (!string.Equals(addressBox.Text, current, StringComparison.OrdinalIgnoreCase))
                    {
                        addressBox.Text = current;
                    }
                }

                UpdateNavButtonState();
            };

            NavigateWebView(webView, addressBox.Text, item);
            UpdateNavButtonState();
        }

        webView.Loaded += InitializeWebView;

        backButton.Click += (_, _) =>
        {
            if (webView.CanGoBack)
            {
                webView.GoBack();
            }
        };

        forwardButton.Click += (_, _) =>
        {
            if (webView.CanGoForward)
            {
                webView.GoForward();
            }
        };

        refreshButton.Click += (_, _) =>
        {
            if (webView.Source is null)
            {
                NavigateWebView(webView, addressBox.Text, item);
            }
            else
            {
                webView.Reload();
            }
        };

        addressBox.KeyDown += (_, args) =>
        {
            if (args.Key == Windows.System.VirtualKey.Enter)
            {
                NavigateWebView(webView, addressBox.Text, item);
                args.Handled = true;
            }
        };

        return layout;
    }

    private static void NavigateWebView(WebView2 webView, string? rawAddress, BoardItemModel item)
    {
        var url = NormalizeWebAddress(rawAddress);
        if (string.IsNullOrWhiteSpace(url) || !Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return;
        }

        item.Content = uri.ToString();
        webView.Source = uri;
    }

    private static string NormalizeWebAddress(string? rawAddress)
    {
        var text = rawAddress?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        if (Uri.TryCreate(text, UriKind.Absolute, out _))
        {
            return text;
        }

        return $"https://{text}";
    }
}
