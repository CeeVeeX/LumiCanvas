using System;
using System.Diagnostics;
using System.IO;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
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
}
