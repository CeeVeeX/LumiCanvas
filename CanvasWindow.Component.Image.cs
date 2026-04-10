using System;
using System.IO;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;

namespace LumiCanvas;

public sealed partial class CanvasWindow
{
    private FrameworkElement BuildImageCard(BoardItemModel item)
    {
        return !string.IsNullOrWhiteSpace(item.SourcePath) && File.Exists(item.SourcePath)
            ? new Image
            {
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Source = CreateOptimizedImageSource(item)
            }
            : CreateMissingMediaHint("图片文件不存在或已被移动");
    }

    private BitmapImage CreateOptimizedImageSource(BoardItemModel item)
    {
        var bitmap = new BitmapImage();
        var decodePixelWidth = (int)Math.Clamp(Math.Ceiling(item.Width * Math.Max(1, _scale) * 1.25), 256, 4096);
        bitmap.DecodePixelWidth = decodePixelWidth;
        bitmap.UriSource = new Uri(item.SourcePath!);
        return bitmap;
    }
}
