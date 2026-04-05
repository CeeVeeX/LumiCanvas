using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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
    private static readonly object WebView2EnvironmentLock = new();
    private static readonly object EmbeddedWebAssetsLock = new();
    private static Task<CoreWebView2Environment>? _webView2EnvironmentTask;
    private static string? _embeddedWebAssetsPath;
    private Dictionary<string, string>? _activeFootnoteDefinitions;
    private Dictionary<string, int>? _activeFootnoteOrder;

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
            editor.LostFocus += MarkdownEditor_LostFocus;

            layout.Children.Add(editor);
            return layout;
        }

        var content = GetMarkdownContent(item);
        if (content is null)
        {
            return CreateMissingMediaHint("Markdown文件不存在或已被移动");
        }

        layout.Children.Add(new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Padding = new Thickness(12),
            Content = BuildMarkdownPreview(content)
        });
        return layout;
    }

    private static string? GetMarkdownContent(BoardItemModel item)
    {
        if (!string.IsNullOrWhiteSpace(item.SourcePath))
        {
            if (!File.Exists(item.SourcePath))
            {
                return null;
            }

            try
            {
                return File.ReadAllText(item.SourcePath);
            }
            catch
            {
                return null;
            }
        }

        return item.Content ?? string.Empty;
    }

    private FrameworkElement BuildMarkdownPreview(string markdown)
    {
        var panel = new StackPanel { Spacing = 8 };
        var rawLines = (markdown ?? string.Empty).Replace("\r\n", "\n").Split('\n');
        var lines = ExtractFootnoteDefinitions(rawLines, out var footnotes);
        _activeFootnoteDefinitions = footnotes;
        _activeFootnoteOrder = [];

        var codeBuffer = string.Empty;
        string? codeLanguage = null;
        string? fenceToken = null;
        var insideCodeBlock = false;

        for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
        {
            var rawLine = lines[lineIndex];
            var line = rawLine ?? string.Empty;
            if (TryGetFenceInfo(line, out var currentFenceToken, out var currentFenceLanguage))
            {
                if (insideCodeBlock)
                {
                    if (!string.Equals(fenceToken, currentFenceToken, StringComparison.Ordinal))
                    {
                        codeBuffer += line + "\n";
                        continue;
                    }

                    panel.Children.Add(CreateCodeBlock(codeBuffer.TrimEnd('\n'), codeLanguage));
                    codeBuffer = string.Empty;
                    codeLanguage = null;
                    fenceToken = null;
                }
                else
                {
                    fenceToken = currentFenceToken;
                    codeLanguage = currentFenceLanguage;
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

            if (lineIndex + 1 < lines.Count && TryCreateSetextHeading(line, lines[lineIndex + 1], out var setextHeading))
            {
                panel.Children.Add(setextHeading);
                lineIndex++;
                continue;
            }

            if (IsMarkdownHorizontalRule(line))
            {
                panel.Children.Add(new Border
                {
                    Height = 1,
                    Margin = new Thickness(0, 6, 0, 6),
                    Background = new SolidColorBrush(ColorHelper.FromArgb(120, ItemBorderBrush.Color.R, ItemBorderBrush.Color.G, ItemBorderBrush.Color.B))
                });
                continue;
            }

            if (TryCreateMarkdownTable(lines, lineIndex, out var tableElement, out var consumedLines))
            {
                panel.Children.Add(tableElement);
                lineIndex += consumedLines - 1;
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
            panel.Children.Add(CreateCodeBlock(codeBuffer.TrimEnd('\n'), codeLanguage));
        }

        AppendRenderedFootnotes(panel);

        if (panel.Children.Count == 0)
        {
            panel.Children.Add(new TextBlock
            {
                Foreground = SecondaryTextBrush,
                Text = "空白笔记",
                FontStyle = FontText.FontStyle.Italic
            });
        }

        _activeFootnoteDefinitions = null;
        _activeFootnoteOrder = null;
        return panel;
    }

    private static bool TryGetFenceInfo(string line, out string fenceToken, out string? language)
    {
        fenceToken = string.Empty;
        language = null;

        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var trimmed = line.TrimStart();
        if (!trimmed.StartsWith("```", StringComparison.Ordinal) && !trimmed.StartsWith("~~~", StringComparison.Ordinal))
        {
            return false;
        }

        var tokenChar = trimmed[0];
        var length = 0;
        while (length < trimmed.Length && trimmed[length] == tokenChar)
        {
            length++;
        }

        if (length < 3)
        {
            return false;
        }

        fenceToken = new string(tokenChar, length);
        language = trimmed.Length > length ? trimmed[length..].Trim() : null;
        return true;
    }

    private static List<string> ExtractFootnoteDefinitions(string[] lines, out Dictionary<string, string> footnotes)
    {
        footnotes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var renderLines = new List<string>();

        var index = 0;
        while (index < lines.Length)
        {
            var line = lines[index] ?? string.Empty;
            if (!TryParseFootnoteDefinitionHeader(line, out var footnoteKey, out var firstValuePart))
            {
                renderLines.Add(line);
                index++;
                continue;
            }

            var valueLines = new List<string> { firstValuePart };
            index++;
            while (index < lines.Length)
            {
                var continuation = lines[index] ?? string.Empty;
                if (continuation.StartsWith("    ", StringComparison.Ordinal) || continuation.StartsWith("\t", StringComparison.Ordinal))
                {
                    valueLines.Add(continuation.TrimStart(' ', '\t'));
                    index++;
                    continue;
                }

                break;
            }

            footnotes[footnoteKey] = string.Join("\n", valueLines).Trim();
        }

        return renderLines;
    }

    private static bool TryParseFootnoteDefinitionHeader(string line, out string key, out string value)
    {
        key = string.Empty;
        value = string.Empty;
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var match = Regex.Match(line, "^\\[\\^([^\\]]+)\\]:\\s*(.*)$");
        if (!match.Success)
        {
            return false;
        }

        key = match.Groups[1].Value.Trim();
        value = match.Groups[2].Value;
        return !string.IsNullOrWhiteSpace(key);
    }

    private void AppendRenderedFootnotes(StackPanel panel)
    {
        if (_activeFootnoteDefinitions is null || _activeFootnoteOrder is null || _activeFootnoteOrder.Count == 0)
        {
            return;
        }

        panel.Children.Add(new Border
        {
            Height = 1,
            Margin = new Thickness(0, 10, 0, 6),
            Background = new SolidColorBrush(ColorHelper.FromArgb(120, ItemBorderBrush.Color.R, ItemBorderBrush.Color.G, ItemBorderBrush.Color.B))
        });

        foreach (var entry in _activeFootnoteOrder.OrderBy(entry => entry.Value))
        {
            if (!_activeFootnoteDefinitions.TryGetValue(entry.Key, out var content))
            {
                continue;
            }

            var item = CreateStyledTextBlock($"[{entry.Value}] {content}", 12, NormalWeight);
            item.Foreground = SecondaryTextBrush;
            panel.Children.Add(item);
        }
    }

    private bool TryCreateSetextHeading(string line, string underlineLine, out FrameworkElement headingElement)
    {
        headingElement = null!;
        if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(underlineLine))
        {
            return false;
        }

        var underline = underlineLine.Trim();
        if (underline.Length < 3)
        {
            return false;
        }

        if (underline.All(ch => ch == '='))
        {
            headingElement = CreateStyledTextBlock(line.Trim(), 28, BoldWeight);
            return true;
        }

        if (underline.All(ch => ch == '-'))
        {
            headingElement = CreateStyledTextBlock(line.Trim(), 22, SemiBoldWeight);
            return true;
        }

        return false;
    }

    private static bool IsMarkdownHorizontalRule(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var compact = line.Replace(" ", string.Empty).Trim();
        if (compact.Length < 3)
        {
            return false;
        }

        return compact.All(ch => ch == '-') ||
               compact.All(ch => ch == '*') ||
               compact.All(ch => ch == '_');
    }

    private bool TryCreateMarkdownTable(IReadOnlyList<string> lines, int startIndex, out FrameworkElement tableElement, out int consumedLines)
    {
        tableElement = null!;
        consumedLines = 0;

        if (startIndex + 1 >= lines.Count)
        {
            return false;
        }

        var headerCells = SplitMarkdownTableRow(lines[startIndex]);
        if (headerCells.Count < 2)
        {
            return false;
        }

        var delimiterCells = SplitMarkdownTableRow(lines[startIndex + 1]);
        if (delimiterCells.Count != headerCells.Count || !delimiterCells.All(IsMarkdownTableDelimiterCell))
        {
            return false;
        }

        var bodyRows = new List<List<string>>();
        var index = startIndex + 2;
        while (index < lines.Count)
        {
            var rowLine = lines[index] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(rowLine))
            {
                break;
            }

            var rowCells = SplitMarkdownTableRow(rowLine);
            if (rowCells.Count != headerCells.Count)
            {
                break;
            }

            bodyRows.Add(rowCells);
            index++;
        }

        tableElement = CreateMarkdownTableElement(headerCells, bodyRows);
        consumedLines = 2 + bodyRows.Count;
        return true;
    }

    private static List<string> SplitMarkdownTableRow(string rowLine)
    {
        if (string.IsNullOrWhiteSpace(rowLine) || !rowLine.Contains('|'))
        {
            return [];
        }

        var trimmed = rowLine.Trim();
        if (trimmed.StartsWith("|", StringComparison.Ordinal))
        {
            trimmed = trimmed[1..];
        }

        if (trimmed.EndsWith("|", StringComparison.Ordinal))
        {
            trimmed = trimmed[..^1];
        }

        return trimmed.Split('|').Select(cell => cell.Trim()).ToList();
    }

    private static bool IsMarkdownTableDelimiterCell(string cell)
    {
        if (string.IsNullOrWhiteSpace(cell))
        {
            return false;
        }

        var core = cell.Trim(':');
        return core.Length >= 3 && core.All(ch => ch == '-');
    }

    private FrameworkElement CreateMarkdownTableElement(IReadOnlyList<string> headerCells, IReadOnlyList<List<string>> bodyRows)
    {
        var grid = new Grid();
        for (var column = 0; column < headerCells.Count; column++)
        {
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        }

        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        for (var row = 0; row < bodyRows.Count; row++)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }

        for (var column = 0; column < headerCells.Count; column++)
        {
            var headerCell = CreateMarkdownTableCell(headerCells[column], true);
            Grid.SetRow(headerCell, 0);
            Grid.SetColumn(headerCell, column);
            grid.Children.Add(headerCell);
        }

        for (var row = 0; row < bodyRows.Count; row++)
        {
            for (var column = 0; column < headerCells.Count; column++)
            {
                var bodyCell = CreateMarkdownTableCell(bodyRows[row][column], false);
                Grid.SetRow(bodyCell, row + 1);
                Grid.SetColumn(bodyCell, column);
                grid.Children.Add(bodyCell);
            }
        }

        return new Border
        {
            BorderBrush = ItemBorderBrush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Child = grid
        };
    }

    private Border CreateMarkdownTableCell(string text, bool isHeader)
    {
        var textBlock = new TextBlock
        {
            Foreground = ItemTextBrush,
            TextWrapping = TextWrapping.WrapWholeWords,
            FontWeight = isHeader ? SemiBoldWeight : NormalWeight
        };
        AddMarkdownInlines(textBlock.Inlines, text);

        return new Border
        {
            BorderBrush = ItemBorderBrush,
            BorderThickness = new Thickness(0.5),
            Padding = new Thickness(8, 6, 8, 6),
            Background = isHeader
                ? new SolidColorBrush(ColorHelper.FromArgb(42, AccentBrush.Color.R, AccentBrush.Color.G, AccentBrush.Color.B))
                : null,
            Child = textBlock
        };
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
        if (line.StartsWith("###### ", StringComparison.Ordinal))
        {
            return CreateStyledTextBlock(line[7..], 14, SemiBoldWeight);
        }

        if (line.StartsWith("##### ", StringComparison.Ordinal))
        {
            return CreateStyledTextBlock(line[6..], 15, SemiBoldWeight);
        }

        if (line.StartsWith("#### ", StringComparison.Ordinal))
        {
            return CreateStyledTextBlock(line[5..], 16, SemiBoldWeight);
        }

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

        if (TryCreateMarkdownBulletLine(line, out var bulletElement))
        {
            return bulletElement;
        }

        if (TryCreateMarkdownTaskListLine(line, out var taskElement))
        {
            return taskElement;
        }

        if (TryCreateMarkdownOrderedLine(line, out var orderedElement))
        {
            return orderedElement;
        }

        if (TryCreateMarkdownQuoteLine(line, out var quoteElement))
        {
            return quoteElement;
        }

        return CreateStyledTextBlock(line, 14, NormalWeight);
    }

    private static (int Index, int IndentUnits) ParseMarkdownLineIndent(string line)
    {
        var index = 0;
        var indentUnits = 0;
        while (index < line.Length)
        {
            if (line[index] == ' ')
            {
                var spaces = 0;
                while (index < line.Length && line[index] == ' ')
                {
                    spaces++;
                    index++;
                }

                indentUnits += spaces / 2;
                continue;
            }

            if (line[index] == '\t')
            {
                indentUnits += 2;
                index++;
                continue;
            }

            break;
        }

        return (index, indentUnits);
    }

    private bool TryCreateMarkdownTaskListLine(string line, out FrameworkElement taskElement)
    {
        taskElement = null!;
        if (string.IsNullOrEmpty(line))
        {
            return false;
        }

        var (index, indentUnits) = ParseMarkdownLineIndent(line);
        if (index + 5 >= line.Length || (line[index] != '-' && line[index] != '*') || line[index + 1] != ' ' || line[index + 2] != '[')
        {
            return false;
        }

        var stateToken = line[index + 3];
        if (line[index + 4] != ']' || line[index + 5] != ' ')
        {
            return false;
        }

        var isChecked = stateToken is 'x' or 'X';
        if (!isChecked && stateToken != ' ')
        {
            return false;
        }

        var content = line[(index + 6)..];
        var marker = isChecked ? "☑" : "☐";
        var taskLine = CreateStyledTextBlock($"{marker} {content}", 14, NormalWeight);
        taskLine.FontFamily = new FontFamily("Segoe UI Symbol");
        taskLine.Margin = new Thickness(indentUnits * 12, 0, 0, 0);
        if (isChecked)
        {
            taskLine.Foreground = SecondaryTextBrush;
        }

        taskElement = taskLine;
        return true;
    }

    private bool TryCreateMarkdownQuoteLine(string line, out FrameworkElement quoteElement)
    {
        quoteElement = null!;
        if (string.IsNullOrEmpty(line))
        {
            return false;
        }

        var (index, indentUnits) = ParseMarkdownLineIndent(line);
        if (index >= line.Length || line[index] != '>')
        {
            return false;
        }

        var level = 0;
        while (index < line.Length && line[index] == '>')
        {
            level++;
            index++;
            if (index < line.Length && line[index] == ' ')
            {
                index++;
            }
        }

        var content = index < line.Length ? line[index..] : string.Empty;
        var quote = CreateStyledTextBlock(content, 14, NormalWeight);
        quote.Foreground = AccentBrush;
        quote.Margin = new Thickness((indentUnits + Math.Max(0, level - 1)) * 12, 0, 0, 0);
        quoteElement = quote;
        return true;
    }

    private bool TryCreateMarkdownBulletLine(string line, out FrameworkElement bulletElement)
    {
        bulletElement = null!;
        if (string.IsNullOrEmpty(line))
        {
            return false;
        }

        var (index, indentUnits) = ParseMarkdownLineIndent(line);

        if (index + 1 >= line.Length || (line[index] != '-' && line[index] != '*') || line[index + 1] != ' ')
        {
            return false;
        }

        var content = line[(index + 2)..];
        var bullet = (indentUnits % 3) switch
        {
            1 => "◦",
            2 => "▪",
            _ => "•"
        };

        var bulletLine = CreateStyledTextBlock($"{bullet} {content}", 14, NormalWeight);
        bulletLine.FontFamily = new FontFamily("Segoe UI Symbol");
        bulletLine.Margin = new Thickness(indentUnits * 12, 0, 0, 0);
        bulletElement = bulletLine;
        return true;
    }

    private bool TryCreateMarkdownOrderedLine(string line, out FrameworkElement orderedElement)
    {
        orderedElement = null!;
        if (string.IsNullOrEmpty(line))
        {
            return false;
        }

        var (index, indentUnits) = ParseMarkdownLineIndent(line);

        var markerStart = index;
        while (index < line.Length && char.IsDigit(line[index]))
        {
            index++;
        }

        if (index == markerStart || index + 1 >= line.Length)
        {
            return false;
        }

        var markerEnd = index;
        var delimiter = line[index];
        if ((delimiter != '.' && delimiter != ')') || line[index + 1] != ' ')
        {
            return false;
        }

        var marker = line[markerStart..(markerEnd + 1)];
        var content = line[(index + 2)..];
        var orderedLine = CreateStyledTextBlock($"{marker} {content}", 14, NormalWeight);
        orderedLine.Margin = new Thickness(indentUnits * 12, 0, 0, 0);
        orderedElement = orderedLine;
        return true;
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

    private Border CreateCodeBlock(string code, string? language)
    {
        var hasLanguage = !string.IsNullOrWhiteSpace(language);
        return new Border
        {
            Padding = new Thickness(12),
            CornerRadius = new CornerRadius(12),
            Background = new SolidColorBrush(ColorHelper.FromArgb(255, 14, 20, 28)),
            Child = new StackPanel
            {
                Spacing = hasLanguage ? 6 : 0,
                Children =
                {
                    new TextBlock
                    {
                        Visibility = hasLanguage ? Visibility.Visible : Visibility.Collapsed,
                        Foreground = SecondaryTextBrush,
                        FontSize = 12,
                        Text = language ?? string.Empty
                    },
                    new TextBlock
                    {
                        FontFamily = new FontFamily("Consolas"),
                        Foreground = new SolidColorBrush(ColorHelper.FromArgb(255, 181, 225, 168)),
                        TextWrapping = TextWrapping.Wrap,
                        Text = code
                    }
                }
            }
        };
    }

    private void AddMarkdownInlines(InlineCollection inlines, string text)
    {
        var index = 0;
        while (index < text.Length)
        {
            var boldIndex = text.IndexOf("**", index, StringComparison.Ordinal);
            var strikeIndex = text.IndexOf("~~", index, StringComparison.Ordinal);
            var codeIndex = text.IndexOf('`', index);
            var linkIndex = text.IndexOf('[', index);
            var footnoteRefIndex = text.IndexOf("[^", index, StringComparison.Ordinal);
            var italicAsteriskIndex = FindSingleMarkerIndex(text, '*', index);
            var italicUnderscoreIndex = FindSingleMarkerIndex(text, '_', index);
            while (linkIndex > 0 && text[linkIndex - 1] == '!')
            {
                linkIndex = text.IndexOf('[', linkIndex + 1);
            }

            var nextIndex = MinPositive(
                MinPositive(
                    MinPositive(MinPositive(MinPositive(MinPositive(boldIndex, strikeIndex), codeIndex), linkIndex), footnoteRefIndex),
                    italicAsteriskIndex),
                italicUnderscoreIndex);

            if (nextIndex < 0)
            {
                AddPlainTextWithAutoLinks(inlines, text[index..]);
                return;
            }

            if (nextIndex > index)
            {
                AddPlainTextWithAutoLinks(inlines, text[index..nextIndex]);
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

            if (nextIndex == footnoteRefIndex)
            {
                var end = text.IndexOf(']', footnoteRefIndex + 2);
                if (end > footnoteRefIndex + 2)
                {
                    var key = text[(footnoteRefIndex + 2)..end].Trim();
                    if (_activeFootnoteDefinitions is not null &&
                        _activeFootnoteDefinitions.ContainsKey(key) &&
                        _activeFootnoteOrder is not null)
                    {
                        if (!_activeFootnoteOrder.TryGetValue(key, out var order))
                        {
                            order = _activeFootnoteOrder.Count + 1;
                            _activeFootnoteOrder[key] = order;
                        }

                        inlines.Add(new Run
                        {
                            Text = $"[{order}]",
                            Foreground = AccentBrush,
                            FontSize = 11
                        });
                        index = end + 1;
                        continue;
                    }
                }
            }

            if (nextIndex == strikeIndex)
            {
                var end = text.IndexOf("~~", strikeIndex + 2, StringComparison.Ordinal);
                if (end > strikeIndex)
                {
                    inlines.Add(new Run
                    {
                        Text = ApplyStrikethrough(text[(strikeIndex + 2)..end]),
                        Foreground = SecondaryTextBrush
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

            if (nextIndex == italicAsteriskIndex || nextIndex == italicUnderscoreIndex)
            {
                var marker = text[nextIndex];
                var end = FindSingleMarkerIndex(text, marker, nextIndex + 1);
                if (end > nextIndex + 1)
                {
                    inlines.Add(new Run
                    {
                        Text = text[(nextIndex + 1)..end],
                        FontStyle = FontText.FontStyle.Italic
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

    private void AddPlainTextWithAutoLinks(InlineCollection inlines, string segment)
    {
        if (string.IsNullOrEmpty(segment))
        {
            return;
        }

        var normalized = UnescapeMarkdownLiterals(segment);
        var matches = Regex.Matches(normalized, "https?://[^\\s<>\"\\)]+|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}", RegexOptions.IgnoreCase);
        if (matches.Count == 0)
        {
            inlines.Add(new Run { Text = normalized });
            return;
        }

        var index = 0;
        foreach (Match match in matches)
        {
            if (!match.Success)
            {
                continue;
            }

            if (match.Index > index)
            {
                inlines.Add(new Run { Text = normalized[index..match.Index] });
            }

            var urlText = match.Value;
            var uriText = urlText.Contains('@') && !urlText.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                ? $"mailto:{urlText}"
                : urlText;
            if (Uri.TryCreate(uriText, UriKind.Absolute, out var uri))
            {
                var hyperlink = new Hyperlink();
                hyperlink.Click += (_, _) => HandleMarkdownHyperlink(uri);
                hyperlink.Inlines.Add(new Run
                {
                    Text = urlText,
                    Foreground = AccentBrush
                });
                inlines.Add(hyperlink);
            }
            else
            {
                inlines.Add(new Run { Text = urlText });
            }

            index = match.Index + match.Length;
        }

        if (index < normalized.Length)
        {
            inlines.Add(new Run { Text = normalized[index..] });
        }
    }

    private static string UnescapeMarkdownLiterals(string value)
    {
        return value
            .Replace("\\*", "*", StringComparison.Ordinal)
            .Replace("\\_", "_", StringComparison.Ordinal)
            .Replace("\\`", "`", StringComparison.Ordinal)
            .Replace("\\[", "[", StringComparison.Ordinal)
            .Replace("\\]", "]", StringComparison.Ordinal)
            .Replace("\\(", "(", StringComparison.Ordinal)
            .Replace("\\)", ")", StringComparison.Ordinal)
            .Replace("\\#", "#", StringComparison.Ordinal)
            .Replace("\\-", "-", StringComparison.Ordinal)
            .Replace("\\+", "+", StringComparison.Ordinal)
            .Replace("\\.", ".", StringComparison.Ordinal)
            .Replace("\\!", "!", StringComparison.Ordinal)
            .Replace("\\~", "~", StringComparison.Ordinal)
            .Replace("\\\\", "\\", StringComparison.Ordinal);
    }

    private static int FindSingleMarkerIndex(string text, char marker, int startIndex)
    {
        for (var index = startIndex; index < text.Length; index++)
        {
            if (text[index] != marker)
            {
                continue;
            }

            if (index > 0 && text[index - 1] == '\\')
            {
                continue;
            }

            var previousIsMarker = index > 0 && text[index - 1] == marker;
            var nextIsMarker = index + 1 < text.Length && text[index + 1] == marker;
            if (previousIsMarker || nextIsMarker)
            {
                continue;
            }

            return index;
        }

        return -1;
    }

    private static string ApplyStrikethrough(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return value;
        }

        return string.Concat(value.Select(ch => $"{ch}\u0336"));
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

    private async void MarkdownDoneButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is BoardItemModel item)
        {
            await TryCommitMarkdownEditorContentAsync(item);
            _markdownHighlightTimer.Stop();
            _pendingHighlightEditor = null;
            item.IsEditing = false;
            RefreshBoardItemView(item);
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
            var content = GetMarkdownContent(item) ?? string.Empty;
            var markdownJson = JsonSerializer.Serialize(content);
            await webView.ExecuteScriptAsync($"window.setMarkdownValue({markdownJson});");

            if (_pendingFocusItemId == item.Id)
            {
                _pendingFocusItemId = null;
                await webView.ExecuteScriptAsync("window.focusMarkdownEditor();");
            }
        }

        var environment = await GetWebView2EnvironmentAsync();
        await editor.EnsureCoreWebView2Async(environment);

        var monacoLoaderUrl = ConfigureMonacoLocalAssets(editor);
        var monacoVsBaseUrl = string.IsNullOrWhiteSpace(monacoLoaderUrl)
            ? null
            : monacoLoaderUrl.Replace("/loader.js", string.Empty, StringComparison.Ordinal);
        editor.NavigationCompleted += OnNavigationCompleted;
        editor.NavigateToString(BuildMonacoEditorHtml(monacoLoaderUrl, monacoVsBaseUrl, IsLightThemeActive()));
    }

    private string? ConfigureMonacoLocalAssets(WebView2 editor)
    {
        if (editor.CoreWebView2 is null)
        {
            return null;
        }

        var webAssetsPath = ResolveWebAssetsPath();
        if (string.IsNullOrWhiteSpace(webAssetsPath))
        {
            App.WriteDiagnostic("CanvasWindow.ConfigureMonacoLocalAssets", new FileNotFoundException("WebAssets not found for Monaco."));
            return null;
        }

        editor.CoreWebView2.SetVirtualHostNameToFolderMapping(
            "appassets",
            webAssetsPath,
            CoreWebView2HostResourceAccessKind.Allow);

        return "https://appassets/Monaco/min/vs/loader.js";
    }

    private static string? ResolveWebAssetsPath()
    {
        if (!string.IsNullOrWhiteSpace(_embeddedWebAssetsPath) &&
            File.Exists(Path.Combine(_embeddedWebAssetsPath, "Monaco", "min", "vs", "loader.js")))
        {
            return _embeddedWebAssetsPath;
        }

        var searchRoots = new[]
        {
            AppContext.BaseDirectory,
            AppDomain.CurrentDomain.BaseDirectory,
            Path.GetDirectoryName(Environment.ProcessPath ?? string.Empty) ?? string.Empty,
            Directory.GetCurrentDirectory()
        }
        .Where(path => !string.IsNullOrWhiteSpace(path))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

        foreach (var root in searchRoots)
        {
            var candidate = Path.Combine(root, "WebAssets");
            if (File.Exists(Path.Combine(candidate, "Monaco", "min", "vs", "loader.js")))
            {
                return candidate;
            }
        }

        var extractedPath = EnsureEmbeddedWebAssetsExtracted();
        if (!string.IsNullOrWhiteSpace(extractedPath) &&
            File.Exists(Path.Combine(extractedPath, "Monaco", "min", "vs", "loader.js")))
        {
            return extractedPath;
        }

        return null;
    }

    private static string? EnsureEmbeddedWebAssetsExtracted()
    {
        lock (EmbeddedWebAssetsLock)
        {
            if (!string.IsNullOrWhiteSpace(_embeddedWebAssetsPath) &&
                File.Exists(Path.Combine(_embeddedWebAssetsPath, "Monaco", "min", "vs", "loader.js")))
            {
                return _embeddedWebAssetsPath;
            }

        const string resourcePrefix = "LumiCanvas.WebAssets/";
        var assembly = typeof(CanvasWindow).Assembly;
        var resources = assembly.GetManifestResourceNames()
            .Where(name => name.StartsWith(resourcePrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (resources.Count == 0)
        {
            return null;
        }

        var targetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LumiCanvas",
            "WebAssets");

        if (File.Exists(Path.Combine(targetRoot, "Monaco", "min", "vs", "loader.js")))
        {
            _embeddedWebAssetsPath = targetRoot;
            return targetRoot;
        }

        try
        {
            Directory.CreateDirectory(targetRoot);
            foreach (var resourceName in resources)
            {
                var relative = resourceName[resourcePrefix.Length..].Replace('/', Path.DirectorySeparatorChar);
                var targetPath = Path.Combine(targetRoot, relative);
                var targetDirectory = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrWhiteSpace(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                using var resourceStream = assembly.GetManifestResourceStream(resourceName);
                if (resourceStream is null)
                {
                    continue;
                }

                if (File.Exists(targetPath))
                {
                    continue;
                }

                using var fileStream = File.Create(targetPath);
                resourceStream.CopyTo(fileStream);
            }

            _embeddedWebAssetsPath = targetRoot;
            return targetRoot;
        }
        catch (Exception ex)
        {
            App.WriteDiagnostic("CanvasWindow.EnsureEmbeddedWebAssetsExtracted", ex);
            return null;
        }
        }
    }

    private void MarkdownWebView_WebMessageReceived(WebView2 sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        if (sender.Tag is not BoardItemModel item)
        {
            return;
        }

        item.Content = NormalizeEditorText(args.TryGetWebMessageAsString());
    }

    private static string BuildMonacoEditorHtml(string? monacoLoaderUrl, string? monacoVsBaseUrl, bool isLightTheme)
    {
        var loaderScriptTag = string.IsNullOrWhiteSpace(monacoLoaderUrl)
            ? string.Empty
            : $"<script src=\"{monacoLoaderUrl}\"></script>";
        var monacoVsBaseUrlJson = JsonSerializer.Serialize(monacoVsBaseUrl ?? string.Empty);
        var fallbackBackground = isLightTheme ? "#f8fbff" : "transparent";
        var fallbackColor = isLightTheme ? "#2f3d4f" : "#d6e0ec";
        var monacoTheme = isLightTheme ? "vs" : "vs-dark";

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
                          background: __FALLBACK_BG__;
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

                        window.getMarkdownValue = function () {
                          if (editor) {
                            return editor.getValue();
                          }

                          return fallback.value || '';
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
                              theme: '__MONACO_THEME__',
                             automaticLayout: true,
                             minimap: { enabled: false },
                             wordWrap: 'on',
                             lineNumbers: 'off',
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
            .Replace("__MONACO_VS_BASE_JSON__", monacoVsBaseUrlJson, StringComparison.Ordinal)
            .Replace("__MONACO_THEME__", monacoTheme, StringComparison.Ordinal)
            .Replace("#d6e0ec", fallbackColor, StringComparison.Ordinal)
            .Replace("__FALLBACK_BG__", fallbackBackground, StringComparison.Ordinal);
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

    private static Task<CoreWebView2Environment> GetWebView2EnvironmentAsync()
    {
        lock (WebView2EnvironmentLock)
        {
            _webView2EnvironmentTask ??= CreateWebView2EnvironmentAsync();
            return _webView2EnvironmentTask;
        }
    }

    private static async Task<CoreWebView2Environment> CreateWebView2EnvironmentAsync()
    {
        var userDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "LumiCanvas",
            "WebView2");

        Directory.CreateDirectory(userDataFolder);
        return await CoreWebView2Environment.CreateWithOptionsAsync(null, userDataFolder, null);
    }
}
