using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.Web.WebView2.Core;
using DocText = Microsoft.UI.Text;
using FontText = Windows.UI.Text;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
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
}
